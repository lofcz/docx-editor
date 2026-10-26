import { expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { buildBlockIndex } from '../blocks/block-index';
import { previewDirectPlan } from '../ai/apply-direct';
import { domainNodeFromAst } from './domain-ast';
import { parseTargetMarkdown } from '../projection/parse-target';
import { planTargetDocument } from './plan-target';

for (const type of ['callout', 'story', 'solution']) {
  test(`${type} targets create native nested content and retain identity in a second wave`, async () => {
    const input = await targetDocumentInput('<w:p><w:r><w:t>Seed.</w:t></w:r></w:p>');
    const attributes =
      type === 'callout'
        ? 'type="warning" icon="🧪" backgroundColor="#fff3cc" opaque="keep"'
        : 'narrator="teacher" opaque="keep"';
    const target = (text: string) =>
      `<${type} ${attributes}>\n\n${text}\n\n- Outer\n  - Nested\n\n</${type}>`;
    const planned = planTargetDocument(input, target('First body.'), {
      allocateBlockId: () => 'owner',
    });
    if (!planned.success) throw Error(planned.error);
    const applied = previewDirectPlan(input.pkg, planned.plan);
    if (!applied.success) throw Error(applied.error);
    const before = buildBlockIndex(applied.pkg).byId.get('owner')!;
    expect(before.payload.type).toBe(type);
    expect(before.payload.attributes).toEqual(
      type === 'callout'
        ? { variant: 'warning', icon: '🧪', backgroundColor: '#fff3cc', opaque: 'keep' }
        : { narrator: 'teacher', opaque: 'keep' }
    );
    const projection = await projectSourcePackage(applied.pkg);
    expect(projection.text).toContain('First body\\.');
    expect(projection.text).toMatch(/\n +[-*] Nested/);
    const changed = planTargetDocument(
      { ...input, pkg: applied.pkg, projection },
      target('Second body.')
    );
    if (!changed.success) throw Error(changed.error);
    expect(changed.plan.operations.some((operation) => operation.kind === 'text')).toBe(true);
    const final = previewDirectPlan(applied.pkg, changed.plan);
    if (!final.success) throw Error(final.error);
    const after = buildBlockIndex(final.pkg).byId.get('owner')!;
    expect(after.controlId).toBe(before.controlId);
    expect(after.payload).toEqual(before.payload);
    const reopened = readOoxmlPackage(writeOoxmlPackage(final.pkg));
    if (!reopened.ok) throw Error(reopened.reason);
    const exported = await projectSourcePackage(reopened.package);
    expect(exported.text).toContain('Second body\\.');
    expect(exported.text).not.toContain('First body\\.');
    expect(exported.text).toMatch(/\n +[-*] Nested/);
    expect(
      canonicalOoxmlFingerprint(reopened.package.parts.get(reopened.package.mainDocumentPart)!)
    ).toContain('Second body.');
  });
}

for (const attributes of [
  'type=" WARNING " variant="tip"',
  'variant=" Warning "',
  'type="" variant="warning"',
]) {
  test(`callout import uses the Plate variant rules: ${attributes}`, () => {
    const parsed = parseTargetMarkdown(`<callout ${attributes} icon="🧪">\n\nBody.\n\n</callout>`);
    if (!parsed.success) throw Error(parsed.error);
    expect(domainNodeFromAst(parsed.target.normalizedAst.children[0]!)).toEqual({
      type: 'callout',
      variant: 'warning',
      icon: '🧪',
    });
  });
}
