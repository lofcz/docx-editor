import { expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { openDocumentForExport } from '@docx-editor.dev/core/export';
import { buildNumberingIndex, resolveStoryListItems } from '@docx-editor.dev/core/layout';
import {
  applyTreeOp,
  bodyStoryRoot,
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  directParagraphProperties,
  ensureListDefinition,
  findNode,
  insertChildren,
  paragraphTextOf,
  readOoxmlPackage,
  readOoxmlPart,
  storyParagraphs,
  withPart,
  writeOoxmlPackage,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
import reference from '../__tests__/fixtures/ai-pairs.json';
import { W } from '../../test/support/document';
import {
  checklistParagraph,
  listDocument,
  listParagraph,
  listPPr,
} from '../../test/support/list-document';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { projectMarkdown } from '../projection/markdown';
import { planListMembership, type ListTarget, LIST_PLAN_LIMITS } from './lists';
import { createEditPlan, decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planSiblingBlocks } from './sibling-blocks';

const revision = { documentId: 'list-test', epoch: 'load-1', revision: 2 };
const children = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
function source(
  body = listParagraph('First', '1') + listParagraph('Second', '1'),
  styles?: string
) {
  const loaded = readOoxmlPackage(listDocument(body, styles));
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = loaded.package,
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  const ids = storyParagraphs(bodyStoryRoot(part)!).map((node) => node.id);
  return { revision, pkg, partName: part.name, ids, part };
}
function planned(input: ReturnType<typeof source>, target: ListTarget, index = 0) {
  const result = planListMembership(input, [{ paragraphId: input.ids[index]!, target }]);
  if (!result.success) throw new Error(result.error);
  return result.plan;
}
const existing = (kind: 'ordered' | 'bullet', numId: string, level = 0): ListTarget => ({
  kind,
  level,
  instance: { kind: 'existing', numId },
});

test('new native instances retain separate marker definitions and starts per level', () => {
  const input = source(listParagraph('Parent') + listParagraph('Child')),
    result = planListMembership(input, [
      {
        paragraphId: input.ids[0]!,
        target: { kind: 'bullet', level: 0, instance: { kind: 'new', key: 'mixed' } },
      },
      {
        paragraphId: input.ids[1]!,
        target: { kind: 'ordered', level: 1, instance: { kind: 'new', key: 'mixed', start: 7 } },
      },
    ]);
  expect(result.success).toBe(true);
  if (result.success)
    expect(result.plan.operations).toMatchObject([
      { after: { kind: 'bullet', level: 0, instance: { key: 'mixed' } } },
      { after: { kind: 'ordered', level: 1, instance: { key: 'mixed', start: 7 } } },
    ]);
});
function apply(pkg: OoxmlPackage, plan: CanonicalEditPlan, tracked = false) {
  let part = pkg.parts.get(pkg.mainDocumentPart)!;
  for (const edit of plan.operations) {
    if (edit.kind !== 'list') throw new Error('Expected list edit');
    const target = edit.after;
    if (target.kind === 'checklist' || edit.before.checkbox)
      throw new Error('This oracle handles native numbering only');
    let numId: string | null = '0';
    if (target.kind !== 'none') {
      if (target.instance.kind === 'existing') numId = target.instance.numId;
      else {
        const made = ensureListDefinition(pkg, target.kind);
        if (!made) throw new Error('Cannot create native numbering');
        pkg = made.pkg;
        numId = made.numId;
      }
    }
    const result = applyTreeOp(part, {
      op: 'setListNumbering',
      paragraphId: edit.paragraphId,
      numId,
      ...(target.kind !== 'none' ? { level: target.level } : {}),
      ...(tracked ? { revision: { author: 'ScioBot', date: '2026-10-19T09:00:00Z' } } : {}),
    });
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  return withPart(pkg, part);
}
function text(node: TargetAstNode): string {
  return typeof node.value === 'string'
    ? node.value
    : Array.isArray(node.children)
      ? (node.children as TargetAstNode[]).map(text).join('')
      : '';
}
function orderedItems(markdown: string) {
  const parsed = parseTargetMarkdown(markdown);
  if (!parsed.success) throw new Error(parsed.error);
  const list = parsed.target.normalizedAst.children.find(
    (node) => node.type === 'list' && node.ordered
  )!;
  return (list.children as TargetAstNode[]).map((item) => (item.children as TargetAstNode[])[0]!);
}

test('C01 item-two insertion emits one structural change and no edits for derived later ordinals', async () => {
  const fixture = reference.cases.find((item) => item.name === 'formatting and lists')!;
  const before = orderedItems(fixture.beforeMarkdown),
    after = orderedItems(fixture.afterMarkdown);
  expect(before.map(text)).toEqual(['First', 'Second']);
  expect(after.map(text)).toEqual(['First', 'Added', 'Second']);
  const input = source(),
    root = bodyStoryRoot(input.part)!;
  const original = canonicalOoxmlFingerprint(input.part);
  const structure = planSiblingBlocks(
    {
      ...input,
      part: input.part,
      parentId: root.id,
      startIndex: 0,
      blocks: before.map((ast, at) => ({ id: input.ids[at]!, ast })),
    },
    after
  );
  if (!structure.success) throw new Error(structure.error);
  expect(structure.plan.operations).toHaveLength(1);
  expect(structure.pairs.map((pair) => ({ id: pair.id, changed: pair.changed }))).toEqual(
    input.ids.map((id) => ({ id, changed: false }))
  );
  const membership = planListMembership(
    input,
    input.ids.map((paragraphId) => ({ paragraphId, target: existing('ordered', '1') }))
  );
  if (!membership.success) throw new Error(membership.error);
  expect(membership.plan.operations).toEqual([]);
  const inserted = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${listParagraph('Added', '1')}</w:body></w:document>`,
    { name: '/test/new-list-item.xml', contentType: 'application/xml' }
  );
  if (!inserted.ok) throw new Error(inserted.reason);
  const node = children(bodyStoryRoot(inserted.part)!)[0]!;
  const landed = insertChildren(input.part, root.id, 1, [node]);
  if (!landed.ok) throw new Error('Native insertion failed');
  for (const id of input.ids)
    expect(findNode(landed.part, id) === findNode(input.part, id)).toBe(true);
  const items = resolveStoryListItems(
    children(bodyStoryRoot(landed.part)!).filter(
      (node): node is OoxmlElement => node.kind !== 'textValue'
    ),
    buildNumberingIndex(input.pkg.parts.get('/word/numbering.xml')!.root),
    undefined
  );
  expect([...items.values()].map((item) => item.ordinal)).toEqual([1, 2, 3]);
  const opened = openDocumentForExport(writeOoxmlPackage(withPart(input.pkg, landed.part)));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    expect(projectMarkdown(await opened.session.layout(), revision).text).toBe(
      '1. First\n\n2. Added\n\n3. Second'
    );
  } finally {
    opened.session.dispose();
  }
  expect(canonicalOoxmlFingerprint(input.part)).toBe(original);
});

test('bullet, ordered, and nesting changes retain text, neighboring IDs, and unrelated properties', () => {
  for (const target of [
    existing('bullet', '2'),
    existing('ordered', '1', 2),
    existing('ordered', '3'),
  ]) {
    const input = source(),
      before = canonicalOoxmlFingerprint(input.part);
    const result = planned(input, target);
    expect(result.operations).toHaveLength(1);
    expect(result.operations[0]).toMatchObject({
      kind: 'list',
      paragraphId: input.ids[0],
      before: { numbering: { numId: '1', level: 0 }, checkbox: null },
      after: target,
    });
    const pkg = apply(input.pkg, result),
      part = pkg.parts.get(pkg.mainDocumentPart)!;
    expect(findNode(part, input.ids[1]!) === findNode(input.part, input.ids[1]!)).toBe(true);
    expect(paragraphTextOf(part, input.ids[0]!)).toBe('First');
    expect(directParagraphProperties(part, input.ids[0]!)).toEqual(
      directParagraphProperties(input.part, input.ids[0]!)
    );
    const accepted = applyTreeOp(apply(input.pkg, result, true).parts.get(pkg.mainDocumentPart)!, {
      op: 'acceptAllRevisions',
    });
    const rejected = applyTreeOp(apply(input.pkg, result, true).parts.get(pkg.mainDocumentPart)!, {
      op: 'rejectAllRevisions',
    });
    if (!accepted.ok || !rejected.ok) throw new Error('Native list review failed');
    expect(canonicalOoxmlFingerprint(accepted.part)).toBe(canonicalOoxmlFingerprint(part));
    expect(canonicalOoxmlFingerprint(rejected.part)).toBe(before);
    expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
    expect(decodeEditPlan(serializeEditPlan(result))).toEqual({ success: true, plan: result });
  }
});

test('new list keys preserve one logical instance without allocating or writing numbering during planning', () => {
  const input = source(listParagraph('A') + listParagraph('B'));
  const before = writeOoxmlPackage(input.pkg);
  const result = planListMembership(
    input,
    input.ids.map((paragraphId) => ({
      paragraphId,
      target: { kind: 'ordered', level: 0, instance: { kind: 'new', key: 'one-list', start: 7 } },
    }))
  );
  if (!result.success) throw new Error(result.error);
  expect(result.plan.operations).toHaveLength(2);
  expect(
    result.plan.operations.every(
      (edit) =>
        edit.kind === 'list' &&
        edit.after.kind === 'ordered' &&
        edit.after.instance.kind === 'new' &&
        edit.after.instance.key === 'one-list' &&
        edit.after.instance.start === 7
    )
  ).toBe(true);
  expect(writeOoxmlPackage(input.pkg)).toEqual(before);
  expect(Object.isFrozen(result.plan) && result.plan.operations.every(Object.isFrozen)).toBe(true);
  expect(decodeEditPlan(serializeEditPlan(result.plan))).toEqual(result);
});

test('checklist creation, checked state, nesting, and removal produce explicit owned marker intents', async () => {
  const plain = source(listParagraph('Task'));
  expect(
    planned(plain, { kind: 'checklist', level: 1, checked: false }).operations[0]
  ).toMatchObject({
    kind: 'list',
    before: { checkbox: null },
    after: { kind: 'checklist', level: 1, checked: false },
  });
  const input = source(checklistParagraph(false)),
    before = canonicalOoxmlFingerprint(input.part);
  for (const target of [
    { kind: 'checklist', level: 0, checked: true },
    { kind: 'checklist', level: 2, checked: false },
    { kind: 'none' },
  ] as const) {
    const plan = planned(input, target);
    expect(plan.operations).toHaveLength(1);
    expect(plan.operations[0]).toMatchObject({
      kind: 'list',
      before: { checkbox: { checked: false, level: 0 } },
      after: target,
    });
    expect(decodeEditPlan(serializeEditPlan(plan))).toEqual({ success: true, plan });
  }
  const plan = planned(input, { kind: 'checklist', level: 0, checked: true });
  const op = plan.operations[0]!;
  if (op.kind !== 'list' || !op.before.checkbox) throw new Error('Missing checkbox');
  const changed = applyTreeOp(input.part, {
    op: 'setContentControlValue',
    controlId: op.before.checkbox.controlId,
    value: { kind: 'checkbox', checked: true },
  });
  if (!changed.ok) throw new Error(changed.reason);
  expect(
    contentControlPropertiesOf(findNode(changed.part, op.before.checkbox.controlId)!).checkbox
      ?.checked
  ).toBe(true);
  // Native symbols occupy one model unit. The export owns their displayed glyph.
  expect(paragraphTextOf(changed.part, input.ids[0]!)?.slice(1)).toBe('Task');
  const opened = openDocumentForExport(writeOoxmlPackage(withPart(input.pkg, changed.part)));
  if (!opened.ok) throw new Error(opened.reason);
  try {
    expect(projectMarkdown(await opened.session.layout(), revision).text).toContain('☒');
  } finally {
    opened.session.dispose();
  }
  expect(paragraphTextOf(input.part, input.ids[0]!)).toBe('☐Task');
  expect(canonicalOoxmlFingerprint(input.part)).toBe(before);
  expect(planned(input, { kind: 'checklist', level: 0, checked: false }).operations).toEqual([]);
  expect(
    planned(source(checklistParagraph(false, 0, 'imported-form')), { kind: 'none' }).operations
  ).toEqual([]);
});

test('inherited numbering remains unchanged, while detaching writes an explicit no-numbering intent', () => {
  const styles = `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="List"><w:pPr>${listPPr('1')}</w:pPr></w:style></w:styles>`;
  const input = source(
    '<w:p><w:pPr><w:pStyle w:val="List"/></w:pPr><w:r><w:t>Inherited</w:t></w:r></w:p>',
    styles
  );
  expect(planned(input, existing('ordered', '1')).operations).toEqual([]);
  const plan = planned(input, { kind: 'none' });
  expect(plan.operations[0]).toMatchObject({
    before: { numbering: { numId: '1', level: 0 } },
    after: { kind: 'none' },
  });
  const pkg = apply(input.pkg, plan);
  expect(
    planned({ ...input, pkg, part: pkg.parts.get(pkg.mainDocumentPart)! }, { kind: 'none' })
      .operations
  ).toEqual([]);
});

test('numbering resolves through the native relationship before an unrelated default-path definition', () => {
  const entries = unzipSync(listDocument(listParagraph('First', '1')));
  entries['word/lists/main.xml'] = entries['word/numbering.xml']!;
  entries['word/numbering.xml'] = strToU8(
    strFromU8(entries['word/numbering.xml']!).replaceAll('w:val="decimal"', 'w:val="bullet"')
  );
  entries['word/_rels/document.xml.rels'] = strToU8(
    strFromU8(entries['word/_rels/document.xml.rels']!).replace(
      'Target="numbering.xml"',
      'Target="lists/main.xml"'
    )
  );
  entries['[Content_Types].xml'] = strToU8(
    strFromU8(entries['[Content_Types].xml']!).replace(
      '</Types>',
      '<Override PartName="/word/lists/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
    )
  );
  const loaded = readOoxmlPackage(zipSync(entries));
  if (!loaded.ok) throw new Error(loaded.reason);
  const pkg = loaded.package,
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  const input = {
    revision,
    pkg,
    part,
    partName: part.name,
    ids: storyParagraphs(bodyStoryRoot(part)!).map((node) => node.id),
  };
  expect(planned(input, existing('ordered', '1')).operations).toEqual([]);
  expect(
    planListMembership(input, [{ paragraphId: input.ids[0]!, target: existing('bullet', '1') }])
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test('checklist ownership follows Proposed inline containers and preserves unrelated form controls', () => {
  const wrapped = (tag: 'ins' | 'del') =>
    checklistParagraph(false)
      .replace('<w:sdt>', `<w:${tag} w:id="5" w:author="Ada"><w:sdt>`)
      .replace('</w:sdt>', `</w:sdt></w:${tag}>`);
  expect(
    planned(source(wrapped('ins')), { kind: 'checklist', checked: false, level: 0 }).operations
  ).toEqual([]);
  expect(planned(source(wrapped('del')), { kind: 'none' }).operations).toEqual([]);
  expect(
    planned(source(wrapped('del')), { kind: 'checklist', checked: true, level: 0 }).operations[0]
  ).toMatchObject({ before: { checkbox: null }, after: { kind: 'checklist', checked: true } });
  const foreign = source(checklistParagraph(false, 0, 'ordinary-form'));
  const plan = planned(foreign, { kind: 'checklist', checked: true, level: 0 });
  expect(plan.operations[0]).toMatchObject({ before: { checkbox: null } });
  const malformed = source(
    checklistParagraph(false)
      .replace('<w14:checkbox', '<w14:notCheckbox')
      .replace('</w14:checkbox>', '</w14:notCheckbox>')
  );
  expect(
    planListMembership(malformed, [{ paragraphId: malformed.ids[0]!, target: { kind: 'none' } }])
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test('no-op and malformed targets cannot create phantom ordinals or partial list plans', () => {
  const input = source();
  expect(planListMembership(input, []).success).toBe(true);
  expect(planned(input, existing('ordered', '1')).operations).toEqual([]);
  for (const target of [
    { ...existing('ordered', '1'), ordinal: 3 },
    { kind: 'ordered', level: 9, instance: { kind: 'existing', numId: '1' } },
    { kind: 'checklist', level: 0, checked: 'yes' },
    { kind: 'none', listStart: 3 },
    { kind: 'ordered', level: 0, instance: { kind: 'new', key: 'x', start: -1 } },
  ])
    expect(
      planListMembership(input, [{ paragraphId: input.ids[0]!, target: target as ListTarget }])
    ).toMatchObject({ success: false, code: 'invalid-input' });
  expect(
    planListMembership(input, [{ paragraphId: input.ids[0]!, target: existing('ordered', '99') }])
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
  expect(
    planListMembership(input, [{ paragraphId: input.ids[0]!, target: existing('bullet', '1') }])
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
  expect(
    planListMembership(
      input,
      input.ids.map((paragraphId, at) => ({
        paragraphId,
        target: {
          kind: at ? 'bullet' : 'ordered',
          level: 0,
          instance: { kind: 'new', key: 'same' },
        },
      }))
    )
  ).toMatchObject({ success: false, code: 'ambiguous-target' });
  expect(
    planListMembership(input, [
      { paragraphId: input.ids[0]!, target: { kind: 'none' } },
      { paragraphId: 'missing', target: { kind: 'none' } },
    ])
  ).toMatchObject({ success: false, code: 'invalid-document' });
  expect(
    planListMembership(
      input,
      Array.from({ length: LIST_PLAN_LIMITS.maxItems + 1 }, () => ({
        paragraphId: 'p',
        target: { kind: 'none' as const },
      }))
    )
  ).toMatchObject({ success: false, code: 'invalid-input' });
  let called = false;
  const target = Object.defineProperty({}, 'kind', {
    enumerable: true,
    get() {
      called = true;
      return 'none';
    },
  }) as ListTarget;
  expect(planListMembership(input, [{ paragraphId: input.ids[0]!, target }])).toMatchObject({
    success: false,
    code: 'invalid-input',
  });
  expect(called).toBe(false);
  const good = planned(input, existing('bullet', '2')).operations[0]!;
  expect(
    createEditPlan(revision, [{ ...good, after: { kind: 'ordered', level: 99 } } as never])
  ).toMatchObject({ success: false, code: 'invalid-input' });
});

test('reading list intents preserves the live native revision, selection, and existing history', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: listDocument(listParagraph('First', '1')),
  });
  try {
    const surface = editor.surface!,
      session = surface.session,
      id = session.paragraphIds()[0]!;
    expect(
      surface.applyAutomationOps(() => [
        { op: 'insertText', paragraphId: id, offset: 5, text: '!' },
      ]).rejected
    ).toBe(false);
    const before = writeOoxmlPackage(session.currentPackage()),
      selection = surface.state().selection;
    const current = session.packageRevision(),
      history = [session.canUndo(), session.canRedo()];
    const result = planListMembership(
      {
        revision: { ...revision, revision: current },
        pkg: session.currentPackage(),
        partName: session.part().name,
      },
      [{ paragraphId: id, target: existing('bullet', '2') }]
    );
    expect(result.success).toBe(true);
    expect(writeOoxmlPackage(session.currentPackage())).toEqual(before);
    expect(session.packageRevision()).toBe(current);
    expect(surface.state().selection).toEqual(selection);
    expect([session.canUndo(), session.canRedo()]).toEqual(history);
  } finally {
    editor.destroy();
  }
});
