import { expect, test } from 'bun:test';
import {
  blankDocumentBytes,
  createDocxEditor,
  type DocxEditorInstance,
} from '@docx-editor.dev/core/editor';
import { paragraphTextOf, readOoxmlPackage, writeOoxmlPackage } from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from './block-index';
import { insertBlock } from './insert';

const URL_ = 'https://cdn.example.com/clip.webm';
const VIDEO = { version: 1, id: 'video-1', type: 'video', attributes: { url: URL_ } } as const;

async function seed() {
  const input = await targetDocumentInput(
    '<w:p><w:r><w:t>Intro</w:t></w:r></w:p><w:p><w:r><w:t>After</w:t></w:r></w:p>'
  );
  return writeOoxmlPackage(input.pkg);
}
const texts = (editor: DocxEditorInstance) => {
  const session = editor.surface!.session;
  return session.paragraphIds().map((id) => paragraphTextOf(session.part(), id));
};
const afterIntro = (editor: DocxEditorInstance) => {
  const part = editor.surface!.session.part();
  const body = part.root.children[0];
  if (!body || body.kind === 'textValue') throw new Error('No body.');
  return body.children[1]!.id;
};
async function markdown(editor: DocxEditorInstance) {
  const read = readOoxmlPackage(new Uint8Array(await editor.save()));
  if (!read.ok) throw new Error(read.reason);
  return (await projectSourcePackage(read.package)).text;
}

test('a linked block writes a Word hyperlink to its source; undo removes it', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: await seed() });
  try {
    const inserted = insertBlock(editor, {
      parentControlId: null,
      beforeId: afterIntro(editor),
      payload: VIDEO,
      label: 'Video',
      paragraphs: ['Video: clip.webm'],
      link: URL_,
    });
    if (!inserted.success) throw new Error(inserted.error);
    expect(texts(editor)).toEqual(['Intro', 'Video: clip.webm', 'After']);
    expect(await markdown(editor)).toBe(`Intro\n\n[Video: clip\\.webm](${URL_})\n\nAfter`);
    const pkg = editor.surface!.session.currentPackage();
    expect(buildBlockIndex(pkg).byId.get('video-1')?.payload.attributes).toEqual({ url: URL_ });
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(editor)).toEqual(['Intro', 'After']);
    // Script and relative targets are refused before any write.
    for (const link of ['javascript:alert(1)', '/admin'])
      expect(
        insertBlock(editor, {
          parentControlId: null,
          beforeId: null,
          payload: { ...VIDEO, id: 'video-2' },
          label: 'Video',
          paragraphs: ['Video'],
          link,
        })
      ).toMatchObject({ success: false, code: 'invalid-input' });
    expect(texts(editor)).toEqual(['Intro', 'After']);
  } finally {
    editor.destroy();
    container.remove();
  }
});

test('the link works in a part that does not declare the relationship prefix', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  // The packaged blank template has no `xmlns:r` on its root.
  const editor = createDocxEditor({ container, document: new Uint8Array(blankDocumentBytes()) });
  try {
    const inserted = insertBlock(editor, {
      parentControlId: null,
      beforeId: null,
      payload: VIDEO,
      label: 'Video',
      paragraphs: ['clip.webm'],
      link: URL_,
    });
    if (!inserted.success) throw new Error(inserted.error);
    expect(await markdown(editor)).toContain(`[clip\\.webm](${URL_})`);
  } finally {
    editor.destroy();
    container.remove();
  }
});

for (const first of [0, 1] as const)
  test(`a linked block converges beside peer typing; receiver ${first}`, async () => {
    const pair = await createTwoEditorFixture({ document: await seed() });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
      pair.pause();
      const inserted = insertBlock(alice, {
        parentControlId: null,
        beforeId: afterIntro(alice),
        payload: VIDEO,
        label: 'Video',
        paragraphs: ['Video: clip.webm'],
        link: URL_,
      });
      if (!inserted.success) throw new Error(inserted.error);
      const after = bob
        .surface!.session.paragraphIds()
        .find((id) => paragraphTextOf(bob.surface!.session.part(), id) === 'After')!;
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: after, offset: 5, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      for (const editor of pair.editors)
        expect(await markdown(editor)).toBe(`Intro\n\n[Video: clip\\.webm](${URL_})\n\nAfter Bob`);
      const saved = await pair.savedDigests();
      expect(saved[0]).toEqual(saved[1]);
    } finally {
      pair.destroy();
    }
  });
