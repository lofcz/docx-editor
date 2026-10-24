import { expect, test } from 'bun:test';
import {
  TreePackageStore,
  applyTreeOp,
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  withPart,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import { propertyBoundaryBody, propertyBoundaryOps } from '../../test/support/property-boundaries';
import { withInitializedReviewGroupStore } from './group-record-store';
import { persistReviewGroups, readReviewGroups } from './groups';
import { propertyDecisionRefusal } from './property-decision-admission';

test('property withdrawal checks package forms protection, beyond per-part validation', () => {
  const loaded = readOoxmlPackage(docxFixture(propertyBoundaryBody));
  if (!loaded.ok) throw new Error(loaded.reason);
  const initialized = withInitializedReviewGroupStore(loaded.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const before = initialized.pkg;
  let part = before.parts.get(before.mainDocumentPart)!;
  const paragraph = part.root.children
    .flatMap((node) => (node.kind === 'textValue' ? [] : node.children))
    .find((node) => node.kind === 'paragraph')!;
  for (const op of propertyBoundaryOps(paragraph.id)) {
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  const pending = persistReviewGroups(before, withPart(before, part));
  const read = readReviewGroups(pending);
  if (!read.success) throw new Error(read.error);
  const groups = read.groups.filter((group) => group.native[0]?.address.author !== 'Peer');
  expect(groups).toHaveLength(2);
  expect(propertyDecisionRefusal(pending, groups)).toBeNull();
  const store = new TreePackageStore(pending, part);
  const protectedResult = store.applyLifecycleOp({ op: 'setDocumentProtection', enforce: true });
  if (!protectedResult.ok) throw new Error(protectedResult.reason);
  const protectedPackage = store.currentPackage();
  const fingerprint = canonicalOoxmlFingerprint(
    protectedPackage.parts.get(pending.mainDocumentPart)!
  );
  // The part itself cannot read settings.xml. It still admits the exact inverse.
  expect(
    applyTreeOp(part, {
      op: 'rejectAllRevisions',
      siteNodeIds: groups.flatMap((group) => group.native.flatMap((member) => member.siteNodeIds)),
    }).ok
  ).toBe(true);
  expect(propertyDecisionRefusal(protectedPackage, groups)).toBe('locked');
  expect(store.currentPackage()).toBe(protectedPackage);
  expect(
    canonicalOoxmlFingerprint(store.currentPackage().parts.get(pending.mainDocumentPart)!)
  ).toEqual(fingerprint);
});
