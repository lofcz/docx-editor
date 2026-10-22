import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  directParagraphProperties,
  findNode,
  paragraphTextOf,
  propertyContainer,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { bodyPart } from '../../test/support/document';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planParagraphFormat, type ParagraphFormatTarget } from './paragraph-format';

const paragraphId = '/word/document.xml#0.0.0';
const revision = { documentId: 'paragraph-format', epoch: 'mount-1', revision: 2 };
const body =
  '<w:p><w:pPr><w:pStyle w:val="Normal"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720" w:right="120" w:hanging="360"/><w:tabs><w:tab w:val="right" w:pos="6000"/></w:tabs><w:pBdr><w:bottom w:val="single" w:sz="4"/></w:pBdr><w:customKeep w:val="yes"/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Keep this text 😀.</w:t></w:r></w:p>';
const source = (part = bodyPart(body)) => ({ revision, part, paragraphId });
function planned(target: ParagraphFormatTarget, part?: OoxmlPart) {
  const result = planParagraphFormat(source(part), target);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
/** Pure native oracle; guarded production execution belongs to A01/R06. */
function apply(part: OoxmlPart, plan: CanonicalEditPlan, tracked = false) {
  const edit = plan.operations[0]!;
  if (edit.kind !== 'format' || edit.target.kind !== 'paragraph')
    throw new Error('Expected paragraph format');
  const bag = new Map(
    directParagraphProperties(part, paragraphId).map((property) => [property.localName, property])
  );
  for (const change of edit.changes) {
    if (!change.after.present) bag.delete(change.name);
    else
      bag.set(change.name, {
        localName: change.name,
        attributes: change.after.value as Record<string, string>,
      });
  }
  const result = applyTreeOp(part, {
    op: 'setParagraphProperties',
    paragraphId,
    properties: [...bag.values()],
    ...(tracked ? { revision: { author: 'ScioBot', date: '2026-10-19T09:00:00Z' } } : {}),
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

test('heading changes retain canonical paragraph/run identity, text, and unrelated native properties', () => {
  const original = source();
  const fingerprint = canonicalOoxmlFingerprint(original.part);
  const changed = planned(
    { properties: [{ localName: 'pStyle', attributes: { val: 'Heading2' } }] },
    original.part
  );
  expect(changed.operations).toEqual([
    {
      kind: 'format',
      partName: original.part.name,
      target: { kind: 'paragraph', paragraphId },
      changes: [
        {
          name: 'pStyle',
          before: { present: true, value: { val: 'Normal' } },
          after: { present: true, value: { val: 'Heading2' } },
        },
      ],
    },
  ]);
  const direct = apply(original.part, changed);
  expect(paragraphTextOf(direct, paragraphId)).toBe('Keep this text 😀.');
  const before = findNode(original.part, paragraphId)!,
    after = findNode(direct, paragraphId)!;
  expect(before.id).toBe(after.id);
  if (before.kind !== 'paragraph' || after.kind !== 'paragraph')
    throw new Error('Missing paragraph');
  const oldRuns = before.children.filter((node) => node.kind === 'run');
  const newRuns = after.children.filter((node) => node.kind === 'run');
  expect(newRuns.map((node) => [node.id, canonicalOoxmlFingerprint(node)])).toEqual(
    oldRuns.map((node) => [node.id, canonicalOoxmlFingerprint(node)])
  );
  expect(canonicalOoxmlFingerprint(original.part)).toBe(fingerprint);
  const tracked = apply(original.part, changed, true);
  const accepted = applyTreeOp(tracked, { op: 'acceptAllRevisions' }),
    rejected = applyTreeOp(tracked, { op: 'rejectAllRevisions' });
  if (!accepted.ok || !rejected.ok) throw new Error('Native paragraph review failed');
  expect(canonicalOoxmlFingerprint(accepted.part)).toBe(canonicalOoxmlFingerprint(direct));
  expect(canonicalOoxmlFingerprint(rejected.part)).toBe(fingerprint);
  expect(decodeEditPlan(serializeEditPlan(changed))).toEqual({ success: true, plan: changed });
});

test('spacing patches keep after/line settings, tab stops, borders, sections, unknown XML, and style', () => {
  const original = source();
  const changed = planned(
    { properties: [{ localName: 'spacing', attributes: { before: '0' } }] },
    original.part
  );
  expect(changed.operations[0]).toMatchObject({
    changes: [
      {
        name: 'spacing',
        before: {
          present: true,
          value: { before: '120', after: '240', line: '360', lineRule: 'auto' },
        },
        after: {
          present: true,
          value: { before: '0', after: '240', line: '360', lineRule: 'auto' },
        },
      },
    ],
  });
  const applied = apply(original.part, changed);
  const oldP = propertyContainer(
    findNode(original.part, paragraphId),
    'paragraphProperties',
    'pPr'
  );
  const newP = propertyContainer(findNode(applied, paragraphId), 'paragraphProperties', 'pPr');
  if (!oldP || !newP || oldP.kind === 'textValue' || newP.kind === 'textValue')
    throw new Error('Missing properties');
  for (const name of ['pStyle', 'ind', 'tabs', 'pBdr', 'sectPr', 'customKeep']) {
    const old = oldP.children.find((node) => node.kind !== 'textValue' && node.localName === name)!;
    const next = newP.children.find(
      (node) => node.kind !== 'textValue' && node.localName === name
    )!;
    expect([next.id, canonicalOoxmlFingerprint(next)]).toEqual([
      old.id,
      canonicalOoxmlFingerprint(old),
    ]);
  }
});

test('indentation patches remove conflicting attributes without clearing independent indentation', () => {
  const changed = planned({
    properties: [{ localName: 'ind', attributes: { hanging: null, firstLine: '240' } }],
  });
  expect(changed.operations[0]).toMatchObject({
    changes: [
      {
        name: 'ind',
        before: { present: true, value: { left: '720', right: '120', hanging: '360' } },
        after: { present: true, value: { left: '720', right: '120', firstLine: '240' } },
      },
    ],
  });
});

test('alignment, direction, and toggle plans retain explicit property presence', () => {
  const changed = planned({
    properties: [
      { localName: 'jc', attributes: { val: 'center' } },
      { localName: 'bidi', attributes: { val: '0' } },
      { localName: 'keepNext' },
    ],
  });
  expect(changed.operations[0]).toMatchObject({
    changes: [
      {
        name: 'jc',
        before: { present: false },
        after: { present: true, value: { val: 'center' } },
      },
      { name: 'bidi', before: { present: false }, after: { present: true, value: { val: '0' } } },
      { name: 'keepNext', before: { present: false }, after: { present: true, value: {} } },
    ],
  });
  expect(planned({ properties: [], remove: ['pStyle'] }).operations[0]).toMatchObject({
    changes: [
      {
        name: 'pStyle',
        before: { present: true, value: { val: 'Normal' } },
        after: { present: false },
      },
    ],
  });
});

test('empty, identical, and already absent targets produce no operations', () => {
  for (const target of [
    { properties: [] },
    { properties: [{ localName: 'pStyle', attributes: { val: 'Normal' } }] },
    { properties: [], remove: ['keepNext'] },
  ])
    expect(planned(target).operations).toEqual([]);
});

test('invalid targets and unsupported structural properties refuse without canonical writes', () => {
  const original = source();
  const before = canonicalOoxmlFingerprint(original.part);
  for (const target of [
    { properties: [{ localName: 'sectPr' }] },
    { properties: [{ localName: 'tabs' }] },
    { properties: [{ localName: 'numPr' }] },
    { properties: [{ localName: 'jc' }, { localName: 'jc' }] },
    { properties: [{ localName: 'pStyle' }], remove: ['pStyle'] },
    { properties: [{ localName: 'pStyle', attributes: { val: 'bad\u0000' } }] },
  ])
    expect(planParagraphFormat(original, target).success).toBe(false);
  expect(
    planParagraphFormat({ ...original, paragraphId: 'missing' }, { properties: [] }).success
  ).toBe(false);
  let read = false;
  const target = Object.defineProperty({}, 'properties', {
    enumerable: true,
    get() {
      read = true;
      return [];
    },
  }) as ParagraphFormatTarget;
  expect(planParagraphFormat(original, target).success).toBe(false);
  expect(read).toBe(false);
  expect(canonicalOoxmlFingerprint(original.part)).toBe(before);
});
