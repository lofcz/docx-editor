import { expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '../package/ooxml-serialize.ts';
import { contentControlsIn } from '../package/content-control-nodes.ts';
import { findNode } from '../package/ooxml-edit.ts';
import { readOoxmlPart, serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { collectRevisionSites } from '../store/tree-op-revisions.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';
import { validateTreeOp } from '../store/tree-op-validate.ts';
import { ADA, apply, part, xml } from './tracked-edit-fixture.ts';
const run = (value: string) =>
  `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${value}</w:t></w:r>`;
const control = (value = run('REF'), lock = 'unlocked') =>
  `<w:sdt><w:sdtPr><w:tag w:val="app:owned"/><w:alias w:val="Reference"/><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent>${value}</w:sdtContent></w:sdt>`;
const ids = (node: import('../package/ooxml-tree.ts').OoxmlNode): string[] => [
  node.id,
  ...(node.kind === 'textValue' ? [] : node.children.flatMap(ids)),
];
for (const linked of [false, true])
  for (const empty of [false, true])
    test(`tracked removal preserves whole native control identity and exact rejection: linked=${linked}/empty=${empty}`, () => {
      const fragment = control(empty ? '' : run(' REF '));
      const before = part(
        `<w:p>${run('Before ')}${linked ? `<w:hyperlink w:anchor="target" w:tooltip="Keep">${fragment}</w:hyperlink>` : fragment}${run(' After')}</w:p>`
      );
      const original = contentControlsIn(before.root)[0]!.node;
      const op = {
        op: 'removeContentControl',
        controlId: original.id,
        keepContent: false,
        revision: ADA,
      } as const;
      expect(validateTreeOp(before, op)).toBeNull();
      const pending = apply(before, op),
        sites = collectRevisionSites(pending);
      expect(sites).toHaveLength(1);
      expect(sites[0]!.refused).toBe(false);
      expect(sites[0]!.node.kind).toBe('revisionDelete');
      const held = findNode(pending, original.id)!;
      expect(ids(held)).toEqual(ids(original));
      expect(xml(pending)).toContain('<w:del ');
      expect(xml(pending)).toContain('<w:sdt>');
      if (!empty)
        expect(xml(pending)).toContain('<w:delText xml:space="preserve"> REF </w:delText>');
      const accepted = apply(pending, {
        op: 'acceptAllRevisions',
        siteNodeIds: [sites[0]!.node.id],
      });
      expect(contentControlsIn(accepted.root)).toEqual([]);
      expect(xml(accepted)).not.toContain('<w:del');
      const rejected = apply(pending, {
        op: 'rejectAllRevisions',
        siteNodeIds: [sites[0]!.node.id],
      });
      expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(before));
      expect(ids(findNode(rejected, original.id)!)).toEqual(ids(original));
      const reopened = readOoxmlPart(serializeOoxmlPart(pending), {
        name: pending.name,
        contentType: pending.contentType,
      });
      if (!reopened.ok) throw Error(reopened.reason);
      const reopenedSites = collectRevisionSites(reopened.part);
      expect(reopenedSites).toHaveLength(1);
      expect(reopenedSites[0]!.refused).toBe(false);
      const reopenedRejected = apply(reopened.part, {
        op: 'rejectAllRevisions',
        siteNodeIds: reopenedSites.map((site) => site.node.id),
      });
      expect(canonicalOoxmlFingerprint(reopenedRejected)).toBe(canonicalOoxmlFingerprint(before));
      expect(canonicalOoxmlFingerprint(before)).not.toBe(canonicalOoxmlFingerprint(pending));
    });

test('tracked removal refuses unwrapping, invalid attribution, blocks, locks, and pending ownership before mutation', () => {
  const variants = [
    { content: control(), patch: { keepContent: true }, reason: 'unsupported' },
    { content: control(), patch: { keepContent: undefined }, reason: 'unsupported' },
    {
      content: control(),
      patch: { revision: { author: '', date: ADA.date } },
      reason: 'invalid-property-value',
    },
    { content: control(run('REF'), 'sdtLocked'), patch: {}, reason: 'locked' },
    { content: control(run('REF'), 'sdtContentLocked'), patch: {}, reason: 'locked' },
    {
      content: `<w:ins w:id="1" w:author="Other">${control()}</w:ins>`,
      patch: {},
      reason: 'indivisible-content',
    },
    {
      content: control(
        `<w:r><w:rPr><w:rPrChange w:id="2" w:author="Other"><w:rPr/></w:rPrChange></w:rPr><w:t>REF</w:t></w:r>`
      ),
      patch: {},
      reason: 'indivisible-content',
    },
  ];
  for (const variant of variants) {
    const before = part(`<w:p>${run('Before ')}${variant.content}${run(' After')}</w:p>`);
    const id = contentControlsIn(before.root)[0]!.node.id,
      saved = xml(before);
    const op = {
      op: 'removeContentControl',
      controlId: id,
      keepContent: false,
      revision: ADA,
      ...variant.patch,
    } as import('../store/tree-op-types.ts').TreeDocOp;
    expect(validateTreeOp(before, op)).toBe(variant.reason);
    expect(applyTreeOp(before, op)).toMatchObject({ ok: false, reason: variant.reason });
    expect(xml(before)).toBe(saved);
  }
  const block = part(control(`<w:p>${run('Block')}</w:p>`));
  expect(
    applyTreeOp(block, {
      op: 'removeContentControl',
      controlId: contentControlsIn(block.root)[0]!.node.id,
      keepContent: false,
      revision: ADA,
    })
  ).toMatchObject({ ok: false, reason: 'unsupported' });
});

test('whole-control deletion preserves complex fields, unknown metadata, and instruction identities', () => {
  const field =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    run(' 7 ') +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const before = part(
    `<w:p>${run('Before ')}${control(field).replace('</w:sdtPr>', '<x:private xmlns:x="urn:test:private" x:value="keep"/></w:sdtPr>')}${run(' After')}</w:p>`
  );
  const original = contentControlsIn(before.root)[0]!.node;
  const pending = apply(before, {
    op: 'removeContentControl',
    controlId: original.id,
    keepContent: false,
    revision: ADA,
  });
  expect(xml(pending)).toContain('<w:delInstrText xml:space="preserve"> PAGE </w:delInstrText>');
  expect(ids(findNode(pending, original.id)!)).toEqual(ids(original));
  const sites = collectRevisionSites(pending);
  expect(sites).toHaveLength(1);
  expect(sites[0]!.refused).toBe(false);
  const rejected = apply(pending, {
    op: 'rejectAllRevisions',
    siteNodeIds: sites.map((site) => site.node.id),
  });
  expect(canonicalOoxmlFingerprint(rejected)).toBe(canonicalOoxmlFingerprint(before));
  expect(ids(findNode(rejected, original.id)!)).toEqual(ids(original));
});
