import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  type OoxmlElement,
  type OoxmlParagraphNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { bodyPart, docxFixture } from '../../test/support/document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';

const paragraph = '<w:p><w:r><w:t>New</w:t></w:r></w:p>';
for (const mode of ['editing', 'suggesting', 'viewing'] as const)
  test(`native sibling automation honors ${mode} mode and keeps history atomic`, () => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: docxFixture('<w:p><w:r><w:t>Kept</w:t></w:r></w:p>'),
      author: 'Host author',
      modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
    });
    try {
      const surface = editor.surface!,
        session = surface.session;
      const anchor = session.paragraphIds()[0]!;
      const parent = session.part().root.children[0] as OoxmlElement;
      const seed = (bodyPart(paragraph).root.children[0] as OoxmlElement)
        .children as readonly OoxmlParagraphNode[];
      const before = canonicalOoxmlFingerprint(session.part());
      const beforeRevision = session.packageRevision();
      expect(editor.setEditingMode(mode).ok).toBe(true);
      const op: TreeDocOp = {
        op: 'insertParagraphs',
        parentId: parent.id,
        beforeId: anchor,
        paragraphs: seed,
      };
      const result = surface.applyAutomationOps(
        () => [op],
        { kind: 'body' },
        undefined,
        false,
        true
      );
      if (mode === 'viewing') {
        expect(result.rejected).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
        expect(session.packageRevision()).toBe(beforeRevision);
        expect(surface.state().canUndo).toBe(false);
      } else {
        expect(result.committed).toBe(true);
        expect(session.paragraphIds()).toHaveLength(2);
        expect(session.paragraphIds()[1]).toBe(anchor);
        const sites = collectRevisionSites(session.part());
        expect(sites).toHaveLength(mode === 'suggesting' ? 2 : 0);
        for (const site of sites)
          expect(site.node.attributes.find((attr) => attr.localName === 'author')?.value).toBe(
            'Host author'
          );
        expect(editor.exec({ type: 'undo' }).ok).toBe(true);
        expect(canonicalOoxmlFingerprint(session.part())).toBe(before);
        expect(surface.state().canUndo).toBe(false);
        expect(editor.exec({ type: 'redo' }).ok).toBe(true);
        expect(session.paragraphIds()).toHaveLength(2);
      }
    } finally {
      editor.destroy();
      container.remove();
    }
  });
