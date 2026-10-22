import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { resetGraphemeBoundary, setGraphemeBoundary } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { applyTextPlan, paragraphPlan, textPlanDocument } from '../../test/support/text-plan';
import { decodeEditPlan, serializeEditPlan } from './plan';
import {
  planParagraphText,
  directParagraphTextOps,
  PARAGRAPH_TEXT_LIMITS,
  type PlannedTextEdit,
} from './paragraph-text';

const revision = { documentId: 'test', epoch: 'load-1', revision: 0 };
const source = (text: string, offset?: number) => ({
  revision,
  partName: '/word/document.xml',
  paragraphId: 'native-p',
  text,
  ...(offset === undefined ? {} : { offset }),
});
function edits(before: string, after: string, offset?: number): readonly PlannedTextEdit[] {
  const result = planParagraphText(source(before, offset), after);
  if (!result.success) throw new Error(result.error);
  return result.plan.operations as readonly PlannedTextEdit[];
}
const effect = (edit: PlannedTextEdit) => ({
  start: edit.start,
  end: edit.end,
  before: edit.before,
  after: edit.after,
});

test('minimal insertion, deletion, and replacement keep exact original UTF-16 offsets', () => {
  expect(edits('cat', 'cats').map(effect)).toEqual([{ start: 3, end: 3, before: '', after: 's' }]);
  expect(edits('cats', 'cat').map(effect)).toEqual([{ start: 3, end: 4, before: 's', after: '' }]);
  expect(edits('Before cat after.', 'Before car after.').map(effect)).toEqual([
    { start: 9, end: 10, before: 't', after: 'r' },
  ]);
  expect(edits('same', 'same')).toEqual([]);
  expect(edits('', '😀').map(effect)).toEqual([{ start: 0, end: 0, before: '', after: '😀' }]);
});

test('disjoint changes keep their shared middle and apply from right to left', () => {
  const planned = edits('abc -- xyz', 'ABC -- XYZ');
  expect(planned.map(effect)).toEqual([
    { start: 7, end: 10, before: 'xyz', after: 'XYZ' },
    { start: 0, end: 3, before: 'abc', after: 'ABC' },
  ]);
  expect(edits('cat', 'car', 20).map(effect)).toEqual([
    { start: 22, end: 23, before: 't', after: 'r' },
  ]);
});

test.each([
  ['A😀B', 'A😃B', { start: 1, end: 3, before: '😀', after: '😃' }],
  ['A👨‍👧B', 'A👩‍👧B', { start: 1, end: 6, before: '👨‍👧', after: '👩‍👧' }],
  ['AéB', 'AéB', { start: 1, end: 3, before: 'é', after: 'é' }],
] as const)('emoji and combining clusters remain whole: %s → %s', (before, after, expected) => {
  expect(edits(before, after).map(effect)).toEqual([expected]);
});

test('generated Unicode changes reconstruct targets without splitting grapheme boundaries', () => {
  const alphabet = ['a', 'b', ' ', '😀', '👩‍👧', 'é', 'Ž', '\n'];
  let seed = 17;
  const next = () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed;
  };
  const segmenter = new Intl.Segmenter('und', { granularity: 'grapheme' });
  for (let sample = 0; sample < 100; sample += 1) {
    const units = Array.from({ length: 20 }, () => alphabet[next() % alphabet.length]!);
    const target = [...units];
    target.splice(next() % target.length, 2, alphabet[next() % alphabet.length]!);
    target.splice(next() % target.length, 0, alphabet[next() % alphabet.length]!);
    const before = units.join(''),
      after = target.join('');
    const boundaries = new Set([
      0,
      before.length,
      ...Array.from(segmenter.segment(before), (entry) => entry.index),
    ]);
    let actual = before,
      right = before.length + 1;
    for (const edit of edits(before, after)) {
      expect(edit.end <= right).toBe(true);
      expect(boundaries.has(edit.start) && boundaries.has(edit.end)).toBe(true);
      expect(before.slice(edit.start, edit.end)).toBe(edit.before);
      actual = actual.slice(0, edit.start) + edit.after + actual.slice(edit.end);
      right = edit.start;
    }
    expect(actual).toBe(after);
  }
});

test('planning uses the single native grapheme strategy and observes strategy changes', () => {
  setGraphemeBoundary({
    segment: (text) => (text ? [{ index: 0, text, utf16From: 0, utf16To: text.length }] : []),
  });
  try {
    expect(edits('cat', 'car').map(effect)).toEqual([
      { start: 0, end: 3, before: 'cat', after: 'car' },
    ]);
    setGraphemeBoundary({
      segment: (text) =>
        text
          .split('')
          .map((text, index) => ({ index, text, utf16From: index, utf16To: index + 1 })),
    });
    expect(planParagraphText(source('😀'), '😃')).toMatchObject({
      success: false,
      code: 'unsupported-operation',
    });
  } finally {
    resetGraphemeBoundary();
  }
  expect(edits('cat', 'car').map(effect)).toEqual([{ start: 2, end: 3, before: 't', after: 'r' }]);
});

test('invalid Unicode, inserted controls, offsets, and oversized text produce no partial plans', () => {
  for (const [before, after] of [
    ['a', '\ud83d'],
    ['\udc00', 'a'],
    ['a', '\0'],
    ['a', '\r'],
    ['a', '\uffff'],
  ])
    expect(planParagraphText(source(before!), after!)).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
  for (const offset of [-1, 0.5, Number.MAX_SAFE_INTEGER])
    expect(planParagraphText(source('a', offset), 'new')).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
  expect(
    planParagraphText(source('a'), 'x'.repeat(PARAGRAPH_TEXT_LIMITS.maxBytes + 1))
  ).toMatchObject({ success: false, code: 'invalid-input' });
  expect(
    planParagraphText(source(''), '\t'.repeat(PARAGRAPH_TEXT_LIMITS.maxInlineAtoms + 1))
  ).toMatchObject({ success: false, code: 'unsupported-operation' });
});

test.each([
  { target: 'Before car after 😀.', count: 1 },
  { target: 'Before newcat after 😀.', count: 1 },
  { target: 'Before cats after 😀.', count: 1 },
  { target: 'Before hippopotamus after 😀.', count: 2 },
])(
  'word edits retain neighboring runs, unknown properties, and history: $target',
  async ({ target, count }) => {
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: textPlanDocument(),
    });
    try {
      const surface = editor.surface!,
        session = surface.session;
      const id = session.paragraphIds()[0]!;
      const paragraph = findNode(session.part(), id)!;
      if (paragraph.kind === 'textValue') throw new Error('Missing paragraph');
      const runs = paragraph.children.filter((node) => node.kind === 'run');
      const preserved = [runs[0]!, runs[2]!].map((node) => ({
        id: node.id,
        fingerprint: canonicalOoxmlFingerprint(node),
      }));
      const before = canonicalOoxmlFingerprint(session.part());
      const plan = paragraphPlan(editor, id, target);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
      expect(plan.operations).toHaveLength(count);
      const decoded = decodeEditPlan(serializeEditPlan(plan));
      expect(decoded).toEqual({ success: true, plan });
      if (!decoded.success) throw new Error(decoded.error);
      expect(applyTextPlan(editor, decoded.plan)).toBe(true);
      expect(paragraphTextOf(session.part(), id)).toBe(target);
      for (const keep of preserved)
        expect(canonicalOoxmlFingerprint(findNode(session.part(), keep.id)!)).toBe(
          keep.fingerprint
        );
      expect(findNode(session.part(), runs[1]!.id)?.kind).toBe('run');
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
      expect(session.canUndo()).toBe(false);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(paragraphTextOf(session.part(), id)).toBe(target);
      const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!reopened.ok) throw new Error(reopened.reason);
      const main = reopened.package.parts.get(reopened.package.mainDocumentPart)!;
      expect(canonicalOoxmlFingerprint(main)).toBe(canonicalOoxmlFingerprint(session.part()));
    } finally {
      editor.destroy();
    }
  }
);

test('native text lowering retains tabs, hard breaks, and hyphen element semantics', () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: textPlanDocument(),
  });
  try {
    const session = editor.surface!.session,
      id = session.paragraphIds()[0]!;
    const target = 'Before cat\tA\nB\u001eC\u001f after 😀.';
    const plan = paragraphPlan(editor, id, target);
    expect(applyTextPlan(editor, plan)).toBe(true);
    expect(paragraphTextOf(session.part(), id)).toBe(target);
    const lowered = plan.operations.flatMap((edit) =>
      directParagraphTextOps(edit as PlannedTextEdit)
    );
    expect(lowered.some((op) => op.op === 'insertTab')).toBe(true);
    expect(lowered.some((op) => op.op === 'insertHardBreak')).toBe(true);
    expect(session.paragraphIds()).toHaveLength(2);
  } finally {
    editor.destroy();
  }
});
