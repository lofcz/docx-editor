import { expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx } from './paginated-surface-fixtures.ts';
import {
  collectReviewItems,
  findNode,
  revisionItemsOf,
  paragraphTextOf,
  canonicalOoxmlFingerprint,
  type OoxmlPart,
} from '../../store/index.ts';

const revision = (kind: 'ins' | 'del', id: number, text: string, author = 'Reviewer') =>
  `<w:${kind} w:id="${id}" w:author="${author}" w:date="2026-10-19T09:00:00Z"><w:r><w:${kind === 'ins' ? 't' : 'delText'}>${text}</w:${kind === 'ins' ? 't' : 'delText'}></w:r></w:${kind}>`;
const module = {
  id: 'leading-replacement-review',
  review: {
    collectReviewItems,
    displayModes: ['all-markup', 'simple-markup', 'proposed', 'original'] as const,
    revisionItemsOfParagraph: (part: OoxmlPart, id: string) => {
      const root = findNode(part, id);
      return root?.kind === 'paragraph' ? revisionItemsOf({ ...part, root }) : [];
    },
  },
};
const mount = (body: string) =>
  createDocxEditor({
    container: document.createElement('div'),
    document: docx(`<w:p>${body}</w:p>`),
    modules: [module],
  });
for (const leading of [false, true])
  for (const action of ['acceptReviewItem', 'rejectReviewItem'] as const)
    test(`${leading ? 'insertion' : 'deletion'}-first replacement keeps one native decision and struck-first range semantics through ${action}`, () => {
      const halves = [revision('del', 1, 'Old'), revision('ins', 2, 'Newer')];
      const editor = mount((leading ? halves.reverse() : halves).join(''));
      try {
        const session = editor.surface!.session,
          id = session.paragraphIds()[0]!;
        const raw = editor.getReviewItems({ pairReplacements: false });
        expect(raw).toHaveLength(2);
        const paired = editor.getReviewItems();
        expect(paired).toHaveLength(1);
        expect(paired[0]!.item).toMatchObject({
          kind: 'revision',
          revisionKind: 'replace',
          text: 'Newer',
          replacedText: 'Old',
          replacedRangeCount: 1,
        });
        expect(editor.getReviewItems()[0]!.key).toBe(paired[0]!.key);
        const placement = paired[0]!;
        const item = placement.item;
        if (item.kind !== 'revision') throw new Error('Missing revision');
        expect(item.ranges.map((range) => [range.start.offset, range.end.offset])).toEqual(
          leading
            ? [
                [5, 8],
                [0, 5],
              ]
            : [
                [0, 3],
                [3, 8],
              ]
        );
        expect(item.addresses.map((address) => address.id)).toEqual(['1', '2']);
        const pending = canonicalOoxmlFingerprint(session.part());
        expect(editor.setActiveReviewItem(placement.key).ok).toBe(true);
        expect(placement.pageIndex).toBe(0);
        expect(editor.surface!.state().selection!.anchor.offset).toBe(leading ? 5 : 0);
        expect(editor[action](placement.key).ok).toBe(true);
        expect(paragraphTextOf(session.part(), id)).toBe(
          action === 'acceptReviewItem' ? 'Newer' : 'Old'
        );
        expect(editor.getReviewItems()).toHaveLength(0);
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(pending);
        expect(editor.getReviewItems()).toHaveLength(1);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(paragraphTextOf(session.part(), id)).toBe(
          action === 'acceptReviewItem' ? 'Newer' : 'Old'
        );
      } finally {
        editor.destroy();
      }
    });

test('pairing preserves deletion-first precedence and keeps different authors independent', () => {
  const ambiguous = mount(
    revision('ins', 3, 'Earlier') + revision('del', 1, 'Old') + revision('ins', 2, 'Newer')
  );
  const foreign = mount(revision('ins', 2, 'Newer', 'Another') + revision('del', 1, 'Old'));
  try {
    expect(ambiguous.getReviewItems()).toHaveLength(2);
    expect(
      ambiguous
        .getReviewItems()
        .map((item) => (item.kind === 'revision' ? [item.revisionKind, item.text] : []))
    ).toEqual([
      ['insert', 'Earlier'],
      ['replace', 'Newer'],
    ]);
    expect(foreign.getReviewItems()).toHaveLength(2);
  } finally {
    ambiguous.destroy();
    foreign.destroy();
  }
});
