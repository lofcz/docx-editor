import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  relationshipsOf,
} from '@docx-editor.dev/core/store';
import { docxFixture } from '../../test/support/document';
import {
  setManualList,
  manualListState,
  setManualListChecked,
  adjustManualListLevel,
  manualListMarker,
  LIST_STYLES,
  toggleManualList,
} from './lists';
const context = { documentId: 'manual-lists', epoch: 'mount' };
const fixture = () =>
  createDocxEditor({
    container: document.createElement('div'),
    document: docxFixture(
      '<w:p><w:r><w:t>Alpha</w:t></w:r></w:p><w:p><w:r><w:t>Beta</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
    ),
  });
for (const style of LIST_STYLES)
  test(`${style} native manual list keeps exact counter style, text, nesting, history and Word reopening`, async () => {
    const editor = fixture();
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      const before = canonicalOoxmlFingerprint(s.session.part());
      const result = setManualList(editor, style, context);
      if (!result.success) throw new Error(result.error);
      const marker = manualListMarker(editor, id)!;
      expect(marker).toBeDefined();
      expect(marker.text).toBe(
        {
          disc: '•',
          circle: '○',
          square: '■',
          decimal: '1.',
          'lower-alpha': 'a.',
          'upper-alpha': 'A.',
          'lower-roman': 'i.',
          'upper-roman': 'I.',
        }[style]
      );
      const source = manualListState(editor, context)!;
      expect(source.numbering?.level).toBe(0);
      expect(adjustManualListLevel(editor, 'increase', context).success).toBe(true);
      expect(manualListState(editor, context)!.numbering?.level).toBe(1);
      expect(manualListMarker(editor, id)!.text).toBe(marker.text);
      expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha');
      const nested = canonicalOoxmlFingerprint(s.session.part()),
        nestedPackage = semanticDigest(s.session.currentPackage().parts.values());
      const read = readOoxmlPackage(new Uint8Array(await editor.save()));
      if (!read.ok) throw new Error(read.reason);
      expect(
        canonicalOoxmlFingerprint(read.package.parts.get(read.package.mainDocumentPart)!)
      ).toEqual(nested);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(manualListState(editor, context)!.numbering?.level).toBe(0);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
      // Native numbering allocations are session-persistent (tree-session.ts:ensureListDefinition).
      // Undo removes membership and authored overrides; redo must restore the full definition.
      const shell = s.session.currentPackage();
      expect(manualListState(editor, context)!.numbering).toBeNull();
      expect(
        relationshipsOf(shell, shell.mainDocumentPart).map((r) => [r.type, r.rawTarget])
      ).toEqual([
        [
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
          'numbering.xml',
        ],
      ]);
      expect(shell.parts.get('/word/numbering.xml')).toBeDefined();
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(nested);
      expect(semanticDigest(s.session.currentPackage().parts.values())).toEqual(nestedPackage);
    } finally {
      editor.destroy();
    }
  });
test('manual checkbox creation, checked state, nesting and conversion retain native text and exact history', async () => {
  const editor = fixture();
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      point = { paragraphId: id, offset: 0 };
    s.setSelection({ anchor: point, head: point });
    const before = canonicalOoxmlFingerprint(s.session.part());
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    const created = canonicalOoxmlFingerprint(s.session.part());
    expect(manualListState(editor, context)!.checkbox?.checked).toBe(false);
    expect(setManualListChecked(editor, true, context).success).toBe(true);
    expect(manualListState(editor, context)!.checkbox?.checked).toBe(true);
    expect(adjustManualListLevel(editor, 'increase', context).success).toBe(true);
    expect(manualListState(editor, context)!.checkbox?.level).toBe(1);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(created);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(s.session.part())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(setManualList(editor, 'decimal', context).success).toBe(true);
    expect(manualListState(editor, context)!.checkbox).toBeNull();
    expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha');
  } finally {
    editor.destroy();
  }
});

for (const style of [...LIST_STYLES, 'todo'] as const)
  test(`native ${style} toggle removes only native list membership`, () => {
    const editor = fixture();
    try {
      const s = editor.surface!,
        id = s.session.paragraphIds()[0]!,
        point = { paragraphId: id, offset: 2 };
      s.setSelection({ anchor: point, head: point });
      expect(toggleManualList(editor, style, context).success).toBe(true);
      expect(toggleManualList(editor, style, context).success).toBe(true);
      expect(manualListState(editor, context)!.numbering).toBeNull();
      expect(manualListState(editor, context)!.checkbox).toBeNull();
      expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha');
    } finally {
      editor.destroy();
    }
  });

test('root list outdent leaves membership and checklist text caret survives marker creation', () => {
  const editor = fixture();
  try {
    const s = editor.surface!,
      id = s.session.paragraphIds()[0]!,
      point = { paragraphId: id, offset: 2 };
    s.setSelection({ anchor: point, head: point });
    expect(setManualList(editor, 'todo', context).success).toBe(true);
    expect(s.state().selection.head.offset).toBe(3);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(s.state().selection.head.offset).toBe(2);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(s.state().selection.head.offset).toBe(3);
    expect(adjustManualListLevel(editor, 'decrease', context).success).toBe(true);
    expect(manualListState(editor, context)!.checkbox).toBeNull();
    expect(paragraphTextOf(s.session.part(), id)).toBe('Alpha');
    expect(s.state().selection.head.offset).toBe(2);
    expect(setManualList(editor, 'decimal', context).success).toBe(true);
    expect(adjustManualListLevel(editor, 'decrease', context).success).toBe(true);
    expect(manualListState(editor, context)!.numbering).toBeNull();
  } finally {
    editor.destroy();
  }
});

test('explicit marker styles cannot mutate existing instances or conflict across levels of one new instance', async () => {
  const { planListMembership } = await import('../diff/lists');
  const { readEditorRevision } = await import('../projection/revision-cache');
  const { isListTarget } = await import('../diff/list-types');
  expect(
    isListTarget({
      kind: 'ordered',
      style: 'square',
      level: 0,
      instance: { kind: 'new', key: 'a' },
    })
  ).toBe(false);
  expect(
    isListTarget({
      kind: 'ordered',
      style: 'decimal',
      level: 0,
      instance: { kind: 'existing', numId: '1' },
    })
  ).toBe(false);
  const editor = fixture();
  try {
    const s = editor.surface!,
      ids = s.session.paragraphIds(),
      source = {
        pkg: s.session.currentPackage(),
        partName: s.session.part().name,
        revision: readEditorRevision(editor, context.documentId, context.epoch)!,
      };
    const result = planListMembership(source, [
      {
        paragraphId: ids[0]!,
        target: {
          kind: 'ordered',
          style: 'decimal',
          level: 0,
          instance: { kind: 'new', key: 'shared' },
        },
      },
      {
        paragraphId: ids[1]!,
        target: {
          kind: 'ordered',
          style: 'upper-roman',
          level: 1,
          instance: { kind: 'new', key: 'shared' },
        },
      },
    ]);
    expect(result).toMatchObject({ success: false, code: 'ambiguous-target' });
  } finally {
    editor.destroy();
  }
});
