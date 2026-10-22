import { expect, test } from 'bun:test';
import {
  collectRevisionSites,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  executeProperties,
  propertyEditor,
  propertyLabel,
  propertyToken,
} from '../../test/support/property-execution';
import { groups } from '../../test/support/review-groups';
import { readReviewGroups, type NativeReviewGroup } from './groups';
import type { OwnedProposalGroup } from './owned-proposal-plan';
import {
  captureOwnedProposalContentProofs,
  verifyOwnedProposalContentProofs,
} from './owned-proposal-content';

const revision = { documentId: 'content-evidence', epoch: 'load-1', revision: 0 };
const grant = (group: NativeReviewGroup): OwnedProposalGroup => ({
  groupId: group.id,
  native: group.native.map((member) => ({
    partName: member.partName,
    addresses: [member.address],
  })),
  properties: group.properties,
});

test('content proofs survive a save/reopen and cover valid canonical nesting beyond JSON input depth', () => {
  const loaded = readOoxmlPackage(
    docxFixture(
      '<w:p><w:ins w:id="7" w:author="Bot">' +
        '<w:smartTag>'.repeat(80) +
        '<w:r><w:t>Owned text</w:t></w:r>' +
        '</w:smartTag>'.repeat(80) +
        '</w:ins></w:p>'
    )
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const read = readReviewGroups(loaded.package);
  if (!read.success) throw new Error(read.error);
  expect(read.groups).toHaveLength(1);
  const ownership = read.groups.map(grant);
  const captured = captureOwnedProposalContentProofs(loaded.package, revision, ownership);
  expect(captured.success).toBe(true);
  if (!captured.success) throw new Error(captured.error);
  expect(captured.contentProofs[0]!.fingerprint).toMatch(/^sha256:[a-f0-9]{64}$/);
  const reopened = readOoxmlPackage(writeOoxmlPackage(loaded.package));
  if (!reopened.ok) throw new Error(reopened.reason);
  expect(
    verifyOwnedProposalContentProofs(reopened.package, revision, ownership, captured.contentProofs)
  ).toBeNull();
});

test('property content proofs preserve unrelated label suffix edits outside owned native revisions', () => {
  const editor = propertyEditor();
  try {
    const proposed = executeProperties(editor);
    if (!proposed.success) throw new Error(proposed.error);
    const session = editor.surface!.session;
    const ownership = groups(editor).map(grant);
    const captured = captureOwnedProposalContentProofs(
      session.currentPackage(),
      propertyToken(editor),
      ownership
    );
    if (!captured.success) throw new Error(captured.error);
    const label = propertyLabel(editor);
    const offset = paragraphTextOf(session.part(), label.id)!.length;
    expect(
      editor.surface!.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: label.id, offset, text: ' peer', bias: 'right' },
      ])
    ).toMatchObject({ committed: true, rejected: false });
    const siteText = collectRevisionSites(session.part()).map((site) => JSON.stringify(site.node));
    expect(siteText.some((text) => text.includes(' peer'))).toBe(false);
    expect(
      verifyOwnedProposalContentProofs(
        session.currentPackage(),
        propertyToken(editor),
        ownership,
        captured.contentProofs
      )
    ).toBeNull();
  } finally {
    editor.destroy();
  }
});

test('content proof validation rejects missing, duplicate, malformed, and changed evidence without writes', () => {
  const loaded = readOoxmlPackage(
    docxFixture('<w:p><w:ins w:id="7" w:author="Bot"><w:r><w:t>Owned</w:t></w:r></w:ins></w:p>')
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  const read = readReviewGroups(loaded.package);
  if (!read.success) throw new Error(read.error);
  const ownership = read.groups.map(grant);
  const captured = captureOwnedProposalContentProofs(loaded.package, revision, ownership);
  if (!captured.success) throw new Error(captured.error);
  const bytes = writeOoxmlPackage(loaded.package);
  expect(verifyOwnedProposalContentProofs(loaded.package, revision, ownership, [])).toMatchObject({
    code: 'ambiguous-target',
  });
  expect(
    verifyOwnedProposalContentProofs(loaded.package, revision, ownership, [
      ...captured.contentProofs,
      ...captured.contentProofs,
    ])
  ).toMatchObject({ code: 'invalid-input' });
  expect(
    verifyOwnedProposalContentProofs(loaded.package, revision, ownership, [
      { ...captured.contentProofs[0]!, fingerprint: 'sha256:invalid' },
    ])
  ).toMatchObject({ code: 'invalid-input' });
  expect(
    verifyOwnedProposalContentProofs(loaded.package, revision, ownership, [
      { ...captured.contentProofs[0]!, fingerprint: `sha256:${'0'.repeat(64)}` },
    ])
  ).toMatchObject({ code: 'ambiguous-target' });
  expect(writeOoxmlPackage(loaded.package)).toEqual(bytes);
});
