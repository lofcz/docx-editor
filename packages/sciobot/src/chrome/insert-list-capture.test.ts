import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { escapeXml, paragraphTextOf } from '@docx-editor.dev/core/store';
import captured from '../__tests__/fixtures/native-commands.json';
import { docxFixture } from '../../test/support/document';
import { insertManualList } from './insert-list';
import { manualListState, manualListMarker } from './lists';
const context = { documentId: 'captured-list-insert', epoch: 'mount' };
for (const style of ['disc', 'decimal', 'todo'] as const)
  test(`native list.insert.${style} matches the frozen N01 value and selection`, () => {
    const reference = captured.cases.find((item) => item.id === `list.insert.${style}`)!;
    const source = reference.before.value as { children: { text: string }[] }[];
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: docxFixture(
        source
          .map(
            (node) =>
              `<w:p><w:r><w:t>${escapeXml(node.children.map((child) => child.text).join(''))}</w:t></w:r></w:p>`
          )
          .join('')
      ),
    });
    try {
      const s = editor.surface!,
        ids = s.session.paragraphIds();
      s.setSelection({
        anchor: {
          paragraphId: ids[reference.before.selection!.anchor.path[0]!]!,
          offset: reference.before.selection!.anchor.offset,
        },
        head: {
          paragraphId: ids[reference.before.selection!.focus.path[0]!]!,
          offset: reference.before.selection!.focus.offset,
        },
      });
      const result = insertManualList(editor, style, context);
      if (!result.success) throw new Error(result.error);
      const actualIds = s.session.paragraphIds();
      const actual = actualIds.map((id) => {
        const state = manualListState(editor, context, id)!,
          text = paragraphTextOf(s.session.part(), id)!;
        return {
          text: state.checkbox ? text.slice(1) : text,
          list: !!state.checkbox || !!state.numbering,
          level: state.checkbox?.level ?? state.numbering?.level ?? null,
        };
      });
      const expected = (
        reference.after.value as {
          children: { text: string }[];
          indent?: number;
          listStyleType?: string;
        }[]
      ).map((node) => ({
        text: node.children.map((child) => child.text).join(''),
        list: !!node.listStyleType,
        level: node.indent === undefined ? null : node.indent - 1,
      }));
      expect(actual).toEqual(expected);
      expect(
        style === 'todo'
          ? manualListState(editor, context, result.paragraphId)?.checkbox?.checked
          : manualListMarker(editor, result.paragraphId)?.text
      ).toBe(style === 'todo' ? false : style === 'disc' ? '•' : '1.');
      const point = (p: { paragraphId: string; offset: number }) => {
        const { paragraphId: id, offset } = p;
        return {
          path: [actualIds.indexOf(id), 0],
          offset: offset - (manualListState(editor, context, id)?.checkbox ? 1 : 0),
        };
      };
      expect({
        anchor: point(s.state().selection.anchor),
        focus: point(s.state().selection.head),
      }).toEqual(reference.after.selection!);
    } finally {
      editor.destroy();
    }
  });
