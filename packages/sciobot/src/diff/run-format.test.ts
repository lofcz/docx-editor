import { expect, test } from 'bun:test';
import {
  AUTHORABLE_RUN_PROPERTIES,
  applyTreeOp,
  authoredProperties,
  canonicalOoxmlFingerprint,
  ensureHyperlinkRelationship,
  findNode,
  hyperlinkTargetOf,
  paragraphTextOf,
  propertyContainer,
  readOoxmlPackage,
  relationshipTargetIn,
  withPart,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { bodyPart, docxFixture } from '../../test/support/document';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import {
  planHyperlinkFormat,
  planRunFormat,
  type RunFormatSource,
  type RunFormatTarget,
} from './run-format';

const revision = { documentId: 'format-test', epoch: 'mount-1', revision: 4 };
const paragraphId = '/word/document.xml#0.0.0';
const body =
  '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:r><w:rPr><w:b/><w:rFonts w:ascii="Calibri" w:eastAsia="MS Gothic" w:cs="Arial" w:hint="eastAsia"/><w:color w:val="123456"/><w:customKeep w:val="untouched"/><w:rStyle w:val="KeepMe"/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> after 😀.</w:t></w:r></w:p>';
const formatSource = (part = bodyPart(body), start = 7, end = 11): RunFormatSource => ({
  revision,
  part,
  paragraphId,
  start,
  end,
  displayMode: 'proposed',
});
function plan(source: RunFormatSource, target: RunFormatTarget) {
  const result = planRunFormat(source, target);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
function children(node: OoxmlNode): readonly OoxmlNode[] {
  return node.kind === 'textValue' ? [] : node.children;
}
function runs(part: OoxmlPart) {
  return children(findNode(part, paragraphId)!).filter((node) => node.kind === 'run');
}
/** Test-only lowering. A02/R06 own guarded production execution and native revision attribution. */
function applyFormatting(part: OoxmlPart, plan: CanonicalEditPlan, tracked = false): OoxmlPart {
  let next = part;
  for (const edit of plan.operations) {
    if (edit.kind !== 'format' || edit.target.kind !== 'run')
      throw new Error('Expected run format');
    const owner = runs(next).find((node) => {
      // These oracle cases change only the existing middle run.
      return children(node).some(
        (child) =>
          child.kind !== 'textValue' &&
          child.localName === 't' &&
          child.children.some((text) => text.kind === 'textValue' && text.value === 'Bold')
      );
    })!;
    const properties = new Map(
      authoredProperties(
        propertyContainer(owner, 'runProperties', 'rPr'),
        AUTHORABLE_RUN_PROPERTIES
      ).map((property) => [property.localName, property])
    );
    for (const change of edit.changes) {
      if (!change.after.present) properties.delete(change.name);
      else
        properties.set(change.name, {
          localName: change.name,
          attributes: change.after.value as Record<string, string>,
        });
    }
    const applied = applyTreeOp(next, {
      op: 'setRunProperties',
      paragraphId: edit.target.paragraphId,
      start: edit.target.start,
      end: edit.target.end,
      properties: [...properties.values()],
      ...(tracked ? { revision: { author: 'ScioBot', date: '2026-10-19T09:00:00Z' } } : {}),
    });
    if (!applied.ok) throw new Error(applied.reason);
    next = applied.part;
  }
  return next;
}

test('bold-to-italic emits only the intended property deltas and preserves native text, run IDs, and unrelated XML', () => {
  const source = formatSource();
  const before = canonicalOoxmlFingerprint(source.part);
  const originalRuns = runs(source.part);
  const changed = plan(source, {
    properties: [{ localName: 'b', attributes: { val: '0' } }, { localName: 'i' }],
  });
  expect(changed.operations).toEqual([
    {
      kind: 'format',
      partName: source.part.name,
      target: { kind: 'run', paragraphId, start: 7, end: 11 },
      changes: [
        {
          name: 'b',
          before: { present: true, value: {} },
          after: { present: true, value: { val: '0' } },
        },
        { name: 'i', before: { present: false }, after: { present: true, value: {} } },
      ],
    },
  ]);
  expect(canonicalOoxmlFingerprint(source.part)).toBe(before);
  const directPlan = decodeEditPlan(serializeEditPlan(changed));
  const trackedPlan = decodeEditPlan(serializeEditPlan(changed));
  expect(directPlan).toEqual(trackedPlan);
  expect(JSON.stringify(changed)).not.toMatch(/children|textValue|root|relationships/);
  const applied = applyFormatting(source.part, changed);
  expect(paragraphTextOf(applied, paragraphId)).toBe('Before Bold after 😀.');
  expect(runs(applied).map((node) => node.id)).toEqual(originalRuns.map((node) => node.id));
  expect(canonicalOoxmlFingerprint(findNode(applied, originalRuns[0]!.id)!)).toBe(
    canonicalOoxmlFingerprint(originalRuns[0]!)
  );
  expect(canonicalOoxmlFingerprint(findNode(applied, originalRuns[2]!.id)!)).toBe(
    canonicalOoxmlFingerprint(originalRuns[2]!)
  );
  const oldProperties = children(propertyContainer(originalRuns[1], 'runProperties', 'rPr')!);
  const newProperties = children(
    propertyContainer(findNode(applied, originalRuns[1]!.id), 'runProperties', 'rPr')!
  );
  for (const name of ['color', 'rFonts', 'customKeep', 'rStyle'])
    expect(
      canonicalOoxmlFingerprint(
        newProperties.find((node) => node.kind !== 'textValue' && node.localName === name)!
      )
    ).toBe(
      canonicalOoxmlFingerprint(
        oldProperties.find((node) => node.kind !== 'textValue' && node.localName === name)!
      )
    );
  const tracked = applyFormatting(source.part, changed, true);
  const accepted = applyTreeOp(tracked, { op: 'acceptAllRevisions' });
  const rejected = applyTreeOp(tracked, { op: 'rejectAllRevisions' });
  if (!accepted.ok || !rejected.ok) throw new Error('Native format decision failed');
  expect(canonicalOoxmlFingerprint(accepted.part)).toBe(canonicalOoxmlFingerprint(applied));
  expect(canonicalOoxmlFingerprint(rejected.part)).toBe(before);
  expect(Object.isFrozen(changed.operations)).toBe(true);
});

test('unchanged properties and empty targets emit no operations; removing a property preserves absence', () => {
  const source = formatSource();
  expect(plan(source, { properties: [] }).operations).toEqual([]);
  expect(plan(source, { properties: [{ localName: 'b' }] }).operations).toEqual([]);
  const removal = plan(source, { properties: [], remove: ['b'] });
  expect(removal.operations[0]).toMatchObject({
    changes: [{ name: 'b', before: { present: true, value: {} }, after: { present: false } }],
  });
  const applied = applyFormatting(source.part, removal);
  const properties = authoredProperties(
    propertyContainer(runs(applied)[1], 'runProperties', 'rPr'),
    AUTHORABLE_RUN_PROPERTIES
  );
  expect(properties.some((property) => property.localName === 'b')).toBe(false);
  expect(properties.some((property) => property.localName === 'color')).toBe(true);
  expect(plan(formatSource(applied), { properties: [], remove: ['b'] }).operations).toEqual([]);
});

test('native font-slot merging retains unrequested scripts and removes the overridden theme reference', () => {
  const source = formatSource(
    bodyPart(body.replace('w:ascii="Calibri"', 'w:ascii="Calibri" w:asciiTheme="minorHAnsi"'))
  );
  const changed = plan(source, {
    properties: [{ localName: 'rFonts', attributes: { ascii: 'Arial' } }],
  });
  expect(changed.operations[0]).toMatchObject({
    changes: [
      {
        name: 'rFonts',
        before: {
          present: true,
          value: {
            ascii: 'Calibri',
            asciiTheme: 'minorHAnsi',
            eastAsia: 'MS Gothic',
            cs: 'Arial',
            hint: 'eastAsia',
          },
        },
        after: {
          present: true,
          value: { ascii: 'Arial', eastAsia: 'MS Gothic', cs: 'Arial', hint: 'eastAsia' },
        },
      },
    ],
  });
});

test('native compound-property merging preserves independent settings and complex-script companions', () => {
  const source = formatSource(
    bodyPart(body.replace('<w:b/>', '<w:b/><w:rtl/><w:u w:val="single" w:color="123456"/>'))
  );
  const changed = plan(source, {
    properties: [
      { localName: 'b', attributes: { val: '0' } },
      { localName: 'u', attributes: { val: 'double' } },
    ],
  });
  expect(changed.operations[0]).toMatchObject({
    changes: expect.arrayContaining([
      { name: 'bCs', before: { present: false }, after: { present: true, value: { val: '0' } } },
      {
        name: 'u',
        before: { present: true, value: { val: 'single', color: '123456' } },
        after: { present: true, value: { val: 'double', color: '123456' } },
      },
    ]),
  });
});

test('partial ranges retain UTF-16 offsets and each run’s original property presence', () => {
  const changed = plan(formatSource(undefined, 5, 15), {
    properties: [{ localName: 'b', attributes: { val: '0' } }],
  });
  expect(changed.operations.map((edit) => edit.kind === 'format' && edit.target)).toEqual([
    { kind: 'run', paragraphId, start: 5, end: 7 },
    { kind: 'run', paragraphId, start: 7, end: 11 },
    { kind: 'run', paragraphId, start: 11, end: 15 },
  ]);
  const states = changed.operations.flatMap((edit) =>
    edit.kind === 'format' ? edit.changes.map((change) => change.before.present) : []
  );
  expect(states).toEqual([false, true, false]);
});

test('native review visibility excludes hidden revision halves without rewriting their properties', () => {
  const part = bodyPart(
    '<w:p><w:del w:id="1" w:author="Alice"><w:r><w:delText>old</w:delText></w:r></w:del><w:ins w:id="2" w:author="Bob"><w:r><w:t>new</w:t></w:r></w:ins></w:p>'
  );
  const source = formatSource(part, 0, 6);
  const target = { properties: [{ localName: 'i' }] };
  const proposed = plan(source, target);
  const original = plan({ ...source, displayMode: 'original' }, target);
  expect(proposed.operations[0]).toMatchObject({ target: { start: 3, end: 6 } });
  expect(original.operations[0]).toMatchObject({ target: { start: 0, end: 3 } });
  expect(proposed.operations).toHaveLength(1);
  expect(original.operations).toHaveLength(1);
});

test('invalid input, split graphemes, missing paragraphs, and field ownership refuse without source mutations', () => {
  const source = formatSource();
  const before = canonicalOoxmlFingerprint(source.part);
  for (const target of [
    { properties: [{ localName: 'b' }, { localName: 'b' }] },
    { properties: [{ localName: 'rStyle' }] },
    { properties: [{ localName: 'b' }], remove: ['b'] },
    { properties: [{ localName: 'b', attributes: { val: '\u0000' } }] },
    { properties: [{ localName: 'b', attributes: { 'x:y': 'bad' } }] },
  ])
    expect(planRunFormat(source, target).success).toBe(false);
  for (const range of [
    { start: -1, end: 5 },
    { start: 0, end: 500 },
    { start: 0, end: 0 },
    { start: 19, end: 20 },
  ])
    expect(
      planRunFormat({ ...source, ...range }, { properties: [{ localName: 'b' }] }).success
    ).toBe(false);
  expect(planRunFormat({ ...source, paragraphId: 'missing' }, { properties: [] }).success).toBe(
    false
  );
  const field = formatSource(
    bodyPart(
      '<w:p><w:fldSimple w:instr="DATE"><w:r><w:t>2026</w:t></w:r><w:r><w:t>10</w:t></w:r></w:fldSimple></w:p>'
    ),
    0,
    1
  );
  expect(planRunFormat(field, { properties: [{ localName: 'b' }] })).toMatchObject({
    success: false,
    code: 'unsupported-operation',
  });
  let read = false;
  const target = Object.defineProperty({}, 'properties', {
    enumerable: true,
    get() {
      read = true;
      return [];
    },
  }) as RunFormatTarget;
  expect(planRunFormat(source, target).success).toBe(false);
  expect(read).toBe(false);
  expect(canonicalOoxmlFingerprint(source.part)).toBe(before);
});

function hyperlinkSource() {
  const bytes = docxFixture(
    '<w:p><w:hyperlink r:id="shared" w:tooltip="Keep this"><w:r><w:rPr><w:b/><w:color w:val="112233"/></w:rPr><w:t>Example</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> and </w:t></w:r><w:hyperlink r:id="shared"><w:r><w:t>Sibling</w:t></w:r></w:hyperlink></w:p>',
    {
      relationships:
        '<Relationship Id="shared" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>',
    }
  );
  const loaded = readOoxmlPackage(bytes);
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = loaded.package,
    partName = pkg.mainDocumentPart;
  const link = children(findNode(pkg.parts.get(partName)!, paragraphId)!).find(
    (node) => node.kind === 'hyperlink'
  )!;
  return { revision, pkg, partName, hyperlinkId: link.id };
}

test('a URL-only change targets one native wrapper and retains shared relationships, text, marks, and tooltip', () => {
  const source = hyperlinkSource();
  const before = [...source.pkg.parts].map(([name, part]) => [
    name,
    canonicalOoxmlFingerprint(part),
  ]);
  const changed = planHyperlinkFormat(source, { url: 'https://example.org' });
  if (!changed.success) throw new Error(changed.error);
  expect(changed.plan.operations).toEqual([
    {
      kind: 'format',
      partName: source.partName,
      target: { kind: 'hyperlink', hyperlinkId: source.hyperlinkId },
      changes: [
        {
          name: 'target',
          before: {
            present: true,
            value: { kind: 'external', authored: 'https://example.com', relationshipId: 'shared' },
          },
          after: { present: true, value: { kind: 'external', authored: 'https://example.org' } },
        },
      ],
    },
  ]);
  expect(
    [...source.pkg.parts].map(([name, part]) => [name, canonicalOoxmlFingerprint(part)])
  ).toEqual(before);
  expect(decodeEditPlan(serializeEditPlan(changed.plan))).toEqual({
    success: true,
    plan: changed.plan,
  });
  expect(planHyperlinkFormat(source, { url: 'https://example.com' })).toMatchObject({
    success: true,
    plan: { operations: [] },
  });
  expect(planHyperlinkFormat(source, {})).toMatchObject({
    success: true,
    plan: { operations: [] },
  });
  const changedTarget = changed.plan.operations[0];
  if (changedTarget?.kind !== 'format' || changedTarget.target.kind !== 'hyperlink')
    throw new Error('Expected hyperlink target');
  const delta = changedTarget.changes[0]!;
  if (
    !delta.after.present ||
    !delta.after.value ||
    typeof delta.after.value !== 'object' ||
    Array.isArray(delta.after.value) ||
    !('authored' in delta.after.value)
  )
    throw new Error('Expected target data');
  const ensured = ensureHyperlinkRelationship(
    source.pkg,
    String(delta.after.value.authored),
    source.partName
  );
  if (!ensured) throw new Error('Native URL mint refused');
  const applied = applyTreeOp(ensured.pkg.parts.get(source.partName)!, {
    op: 'setHyperlinkTarget',
    linkId: changedTarget.target.hyperlinkId,
    relationshipId: ensured.relationshipId,
    tooltip: 'Keep this',
  });
  if (!applied.ok) throw new Error(applied.reason);
  const next = withPart(ensured.pkg, applied.part);
  const links = children(findNode(applied.part, paragraphId)!).filter(
    (node) => node.kind === 'hyperlink'
  );
  expect(
    links.map(
      (link) =>
        hyperlinkTargetOf(link, (id) => relationshipTargetIn(next, source.partName, id)).authored
    )
  ).toEqual(['https://example.org', 'https://example.com']);
  expect(paragraphTextOf(applied.part, paragraphId)).toBe('Example and Sibling');
  for (const link of links) {
    const original = findNode(source.pkg.parts.get(source.partName)!, link.id)!;
    expect(children(link).map(canonicalOoxmlFingerprint)).toEqual(
      children(original).map(canonicalOoxmlFingerprint)
    );
  }
});

test('link tooltips distinguish absence from removal, while internal targets retain bookmark text', () => {
  const source = hyperlinkSource();
  expect(planHyperlinkFormat(source, { tooltip: null })).toMatchObject({
    success: true,
    plan: {
      operations: [
        {
          changes: [
            {
              name: 'tooltip',
              before: { present: true, value: 'Keep this' },
              after: { present: false },
            },
          ],
        },
      ],
    },
  });
  expect(planHyperlinkFormat(source, { tooltip: 'Keep this' })).toMatchObject({
    success: true,
    plan: { operations: [] },
  });
  expect(planHyperlinkFormat(source, { anchor: 'bookmark_1' })).toMatchObject({
    success: true,
    plan: {
      operations: [
        {
          changes: [
            {
              name: 'target',
              after: { present: true, value: { kind: 'internal', authored: 'bookmark_1' } },
            },
          ],
        },
      ],
    },
  });
});

test('unsafe URLs and invalid hyperlink targets refuse without minting relationships', () => {
  const source = hyperlinkSource();
  const before = JSON.stringify(source.pkg.externalTargets);
  for (const url of [
    'javascript:alert(1)',
    '/app/delete',
    'data:text/html,bad',
    'https://example.com/' + 'x'.repeat(2048),
  ])
    expect(planHyperlinkFormat(source, { url }).success).toBe(false);
  for (const target of [
    { url: 'https://example.org', anchor: 'bookmark' },
    { tooltip: '\u0000' },
    { anchor: '' },
  ])
    expect(planHyperlinkFormat(source, target).success).toBe(false);
  expect(planHyperlinkFormat({ ...source, hyperlinkId: 'missing' }, {}).success).toBe(false);
  expect(JSON.stringify(source.pkg.externalTargets)).toBe(before);
});
