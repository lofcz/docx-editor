import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { targetDocumentInput } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from './block-index';
import { readColumnGeometry } from './column-geometry';
import { turnIntoColumns } from './turn-into-columns';

async function fixture() {
  return writeOoxmlPackage(
    (
      await targetDocumentInput(
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Selected</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
      )
    ).pkg
  );
}
const select = (editor: DocxEditorInstance) => {
  const surface = editor.surface!;
  const id = surface.session
    .paragraphIds()
    .find((id) => paragraphTextOf(surface.session.part(), id)?.startsWith('Selected'))!;
  const point = { paragraphId: id, offset: 8 };
  surface.setSelection({ anchor: point, head: point });
  return id;
};
const turn = (editor: DocxEditorInstance, count = 3) => {
  const result = turnIntoColumns(editor, {
    count,
    expectedRevision: editor.surface!.session.packageRevision(),
  });
  if (!result.success) throw Error(result.error);
};
const groups = (editor: DocxEditorInstance) => {
  const pkg = editor.surface!.session.currentPackage();
  return [...buildBlockIndex(pkg).byId.values()]
    .filter((block) => block.payload.type === 'column_group')
    .map((group) => readColumnGeometry(pkg, group));
};

test('turn into columns preserves the selected native paragraph, history, and saved structure', async () => {
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await fixture(),
  });
  try {
    const id = select(editor),
      surface = editor.surface!;
    const fingerprint = canonicalOoxmlFingerprint(findNode(surface.session.part(), id)!);
    const before = semanticDigest(surface.session.currentPackage().parts.values());
    turn(editor);
    expect(groups(editor)[0]!.columns).toHaveLength(3);
    expect(canonicalOoxmlFingerprint(findNode(surface.session.part(), id)!)).toBe(fingerprint);
    expect(surface.state().selection.head).toEqual({ paragraphId: id, offset: 0 });
    expect(caretAt(surface.layout(), { paragraphId: id, offset: 0 })).not.toBeNull();
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(semanticDigest(surface.session.currentPackage().parts.values())).toEqual(before);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    select(editor);
    turn(editor, 2);
    expect(groups(editor)).toHaveLength(1);
    expect(groups(editor)[0]!.columns).toHaveLength(2);
    const saved = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!saved.ok) throw Error(saved.reason);
    expect(semanticDigest(saved.package.parts.values())).toEqual(
      semanticDigest(surface.session.currentPackage().parts.values())
    );
  } finally {
    editor.destroy();
  }
});

for (const first of [0, 1] as const)
  for (const action of ['type', 'delete', 'wrap'] as const)
    test(`column wrapping retains concurrent ${action} and history; receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({ document: await fixture(), clientIds: [41, 73] });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
        const id = select(alice);
        select(bob);
        pair.pause();
        turn(alice);
        if (action === 'type') bob.surface!.type(' Bob');
        else if (action === 'delete')
          expect(
            bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).committed
          ).toBe(true);
        else turn(bob);
        pair.reconnect(first);
        await pair.settle();
        const check = async () => {
          for (const editor of pair.editors) {
            for (const group of groups(editor)) expect(group).not.toBeNull();
            const paragraph = findNode(editor.surface!.session.part(), id);
            if (action === 'delete') expect(paragraph).toBeNull();
            else
              expect(paragraphTextOf(editor.surface!.session.part(), id)).toBe(
                action === 'type' ? 'Selected Bob' : 'Selected'
              );
            expect(
              editor.surface!.session.paragraphIds().filter((value) => value === id)
            ).toHaveLength(action === 'delete' ? 0 : 1);
          }
          const saved = await pair.savedDigests();
          expect(saved[0]).toEqual(saved[1]);
        };
        await check();
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        await check();
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        await check();
      } finally {
        pair.destroy();
      }
    });

test('column wrapping refuses protected content, stale selections, and unsupported modes', async () => {
  for (const lock of ['sdtLocked', 'contentLocked', 'sdtContentLocked']) {
    const input = await targetDocumentInput(
      `<w:sdt><w:sdtPr><w:lock w:val="${lock}"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Selected</w:t></w:r></w:p></w:sdtContent></w:sdt>`
    );
    const editor = createDocxEditor({
      container: document.createElement('div'),
      document: writeOoxmlPackage(input.pkg),
    });
    try {
      select(editor);
      const surface = editor.surface!,
        before = surface.session.currentPackage();
      expect(
        turnIntoColumns(editor, { count: 3, expectedRevision: surface.session.packageRevision() })
      ).toMatchObject({ success: false, code: 'unsupported-operation' });
      expect(surface.session.currentPackage()).toBe(before);
    } finally {
      editor.destroy();
    }
  }
  const editor = createDocxEditor({
    container: document.createElement('div'),
    document: await fixture(),
  });
  try {
    select(editor);
    const surface = editor.surface!,
      before = surface.session.currentPackage();
    const request = { count: 3, expectedRevision: surface.session.packageRevision() };
    expect(turnIntoColumns(editor, { ...request, expectedRevision: -1 })).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    for (const mode of ['suggest', 'view'] as const) {
      surface.setEditingMode(mode);
      expect(turnIntoColumns(editor, request)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(surface.editingMode()).toBe(mode);
      expect(surface.session.currentPackage()).toBe(before);
    }
  } finally {
    editor.destroy();
  }
});
