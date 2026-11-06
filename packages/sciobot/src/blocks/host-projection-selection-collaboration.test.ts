import { expect, test } from 'bun:test';
import { bodyStoryRoot, storyParagraphs } from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { zipDocument } from '../../../pro/src/collaboration/__tests__/document-peer-support';

test('a host projection keeps focus while canonical selection reaches its peer after reattach', async () => {
  const pair = await createTwoEditorFixture({
    document: zipDocument(
      '<w:p><w:r><w:t>First line</w:t></w:r></w:p><w:p><w:r><w:t>Second line</w:t></w:r></w:p>'
    ),
  });
  const [alice, bob] = pair.editors;
  const moved = document.createElement('div');
  document.body.append(moved);
  try {
    const source = (editor: typeof alice) => ({
      paragraphs: storyParagraphs(bodyStoryRoot(editor!.surface!.session.part())!),
    });
    for (const reattach of [false, true]) {
      if (reattach) {
        alice!.detach();
        alice!.attach(moved);
      }
      const surface = alice!.surface!;
      const field = document.createElement('div');
      field.dataset.docxHostUi = '';
      field.contentEditable = 'true';
      const text = document.createTextNode('native source projection');
      field.append(text);
      surface.hostElement().append(field);
      field.focus();
      const range = document.createRange();
      range.setStart(text, 3);
      range.setEnd(text, 9);
      const dom = document.getSelection()!;
      dom.removeAllRanges();
      dom.addRange(range);
      const native = source(alice);
      const beforeRevision = surface.session.packageRevision();
      const selection = {
        anchor: { paragraphId: native.paragraphs[0]!.id, offset: 1 },
        head: { paragraphId: native.paragraphs[1]!.id, offset: reattach ? 4 : 2 },
      };
      surface.setSelection(selection, { preserveDomSelection: true });
      await pair.settle();
      expect(surface.state().selection).toEqual(selection);
      expect(document.activeElement).toBe(field);
      expect(dom.anchorNode === text).toBe(true);
      expect(dom.anchorOffset).toBe(3);
      expect(dom.focusNode === text).toBe(true);
      expect(dom.focusOffset).toBe(9);
      expect(surface.session.packageRevision()).toBe(beforeRevision);
      const peer = bob!.surface!.collaborationSession()!.remoteSelections();
      expect(peer).toHaveLength(1);
      const other = source(bob);
      expect(peer[0]!.anchor.nodeId).toBe(other.paragraphs[0]!.id);
      expect(peer[0]!.anchor.offset).toBe(1);
      expect(peer[0]!.head.nodeId).toBe(other.paragraphs[1]!.id);
      expect(peer[0]!.head.offset).toBe(reattach ? 4 : 2);
      field.remove();
    }
  } finally {
    pair.destroy();
    moved.remove();
  }
});
