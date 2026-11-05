// The test transport uses the EigenPal Pro License.
import { expect, test } from 'bun:test';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt, paragraphFragmentsOf, selectionRects } from '@docx-editor.dev/core/layout';
import {
  paragraphTextOf,
  readOoxmlPackage,
  semanticDigest,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import { prepareNativeDocument } from '../native/prepare-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { setManualList, setManualListChecked, manualListState } from './lists';
const context = { documentId: 'task-style-peers', epoch: 'mount' };
const bytes = () => {
  const result = prepareNativeDocument(
    docxFixture(
      '<w:p><w:r><w:t>Task</w:t></w:r></w:p><w:p><w:r><w:t>Other</w:t></w:r></w:p><w:p><w:r><w:t>Tail</w:t></w:r></w:p>'
    )
  );
  if (!result.success) throw new Error(result.error);
  return result.document;
};
const style = (editor: ReturnType<typeof createDocxEditor>, text: string) =>
  editor
    .surface!.layout()
    .pages.flatMap((p) => paragraphFragmentsOf(p))
    .flatMap((p) => p.lines.flatMap((l) => l.spans))
    .find((s) => s.text.includes(text))!.style;
for (const mode of ['editing', 'suggesting'] as const)
  test(`shared ${mode} completion needs no task style infrastructure`, async () => {
    const pair = await createTwoEditorFixture({
      document: docxFixture('<w:p><w:r><w:t>Task</w:t></w:r></w:p>'),
      documentId: context.documentId,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], author: 'Alice' },
    });
    try {
      const alice = pair.editors[0]!;
      expect(setManualList(alice, 'todo', context).success).toBe(true);
      await pair.settle();
      expect(alice.exec({ type: 'setEditingMode', mode }).ok).toBe(true);
      const surface = alice.surface!,
        before = stylesPartOf(surface.session.currentPackage())?.root,
        revision = surface.session.packageRevision(),
        selection = surface.state().selection;
      expect(setManualListChecked(alice, true, context).success).toBe(true);
      expect(stylesPartOf(surface.session.currentPackage())?.root).toBe(before);
      expect(surface.session.packageRevision()).toBeGreaterThan(revision);
      expect(surface.state().selection).toEqual(selection);
      await pair.settle();
      for (const editor of pair.editors) expect(style(editor, 'Task').strike).toBe(true);
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  }, 120_000);
for (const receiver of [0, 1] as const)
  test(`completed tasks retain peer text, geometry, history, deletion recovery, and Word styles: receiver ${receiver}`, async () => {
    const pair = await createTwoEditorFixture({
      document: bytes(),
      documentId: context.documentId,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
    });
    try {
      const [alice, bob] = pair.editors,
        a = alice!.surface!,
        b = bob!.surface!,
        id = a.session.paragraphIds()[0]!;
      const select = (surface: typeof a, offset: number) => {
        const p = { paragraphId: id, offset };
        surface.setSelection({ anchor: p, head: p });
      };
      select(a, 2);
      expect(setManualList(alice!, 'todo', context).success).toBe(true);
      await pair.settle();
      pair.pause();
      expect(setManualListChecked(alice!, true, context).success).toBe(true);
      select(b, paragraphTextOf(b.session.part(), id)!.length);
      b.insertPlainText(' Peer');
      b.flushPendingInput();
      pair.reconnect(receiver);
      await pair.settle();
      for (const e of pair.editors) {
        expect(style(e, 'Task').strike).toBe(true);
        expect(style(e, 'Peer').strike).toBe(true);
      }
      const selection = {
        anchor: { paragraphId: id, offset: 2 },
        head: { paragraphId: id, offset: 7 },
      };
      b.setSelection(selection);
      await pair.settle();
      expect(
        pair.rooms[0]!.session.remoteSelections().find((s) => s.actorId === 'bob')?.head.offset
      ).toBe(7);
      expect(caretAt(a.layout(), selection.head)).toEqual(caretAt(b.layout(), selection.head));
      expect(selectionRects(a.layout(), selection, a.session.paragraphIds())).toEqual(
        selectionRects(b.layout(), selection, b.session.paragraphIds())
      );
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      for (const e of pair.editors) {
        expect(style(e, 'Task').strike).toBe(false);
        expect(paragraphTextOf(e.surface!.session.part(), id)).toContain('Peer');
      }
      expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      pair.pause();
      expect(a.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: id }]).committed).toBe(true);
      select(b, paragraphTextOf(b.session.part(), id)!.length);
      b.insertPlainText(' kept');
      b.flushPendingInput();
      pair.reconnect(receiver === 0 ? 1 : 0);
      await pair.settle();
      expect(paragraphTextOf(a.session.part(), id)).toBeNull();
      expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      const saved = [];
      for (const e of pair.editors) {
        expect(paragraphTextOf(e.surface!.session.part(), id)).toContain('Peer kept');
        expect(style(e, 'Task').strike).toBe(true);
        const savedBytes = new Uint8Array(await e.save());
        const read = readOoxmlPackage(savedBytes);
        if (!read.ok) throw new Error(read.reason);
        saved.push(semanticDigest(read.package.parts.values()));
        const reopened = createDocxEditor({
          container: document.createElement('div'),
          document: savedBytes,
          modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
        });
        try {
          expect(style(reopened, 'Task').strike).toBe(true);
          expect(style(reopened, 'Peer').strike).toBe(true);
        } finally {
          reopened.destroy();
        }
      }
      expect(saved[0]).toEqual(saved[1]);
      pair.ownership.forEach((owner) => owner.assertHealthy());
    } finally {
      pair.destroy();
    }
  }, 120_000);

test('one participant undo cannot remove another task completion appearance', async () => {
  const pair = await createTwoEditorFixture({
    document: bytes(),
    documentId: context.documentId,
    ownedAttributes: true,
    editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] },
  });
  try {
    const [alice, bob] = pair.editors,
      a = alice!.surface!,
      b = bob!.surface!;
    const [first, second] = a.session.paragraphIds();
    const set = (surface: typeof a, id: string) => {
      const p = { paragraphId: id, offset: 1 };
      surface.setSelection({ anchor: p, head: p });
    };
    set(a, first!);
    expect(setManualList(alice!, 'todo', context).success).toBe(true);
    await pair.settle();
    set(b, second!);
    expect(setManualList(bob!, 'todo', context).success).toBe(true);
    await pair.settle();
    expect(setManualListChecked(alice!, true, context).success).toBe(true);
    await pair.settle();
    expect(setManualListChecked(bob!, true, context).success).toBe(true);
    await pair.settle();
    expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
    await pair.settle();
    for (const e of pair.editors) {
      expect(style(e, 'Task').strike).toBe(false);
      expect(style(e, 'Other').strike).toBe(true);
    }
  } finally {
    pair.destroy();
  }
}, 120_000);

for (const action of ['accept', 'reject'] as const)
  test(`suggested task completion derives appearance from checkbox through ${action}`, async () => {
    const pair = await createTwoEditorFixture({
      document: bytes(),
      documentId: context.documentId,
      ownedAttributes: true,
      editor: { modules: [SCIOBOT_NATIVE_REVIEW_MODULE], author: 'Alice' },
    });
    try {
      const [alice, bob] = pair.editors;
      expect(setManualList(alice!, 'todo', context).success).toBe(true);
      await pair.settle();
      expect(alice!.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
      expect(setManualListChecked(alice!, true, context).success).toBe(true);
      await pair.settle();
      expect(style(alice!, 'Task').strike).toBe(true);
      expect(bob!.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);
      await pair.settle();
      for (const e of pair.editors) {
        expect(style(e, 'Task').strike).toBe(action === 'accept');
        expect(manualListState(e, context)!.checkbox!.checked).toBe(action === 'accept');
      }
      expect((await pair.savedDigests())[0]).toEqual((await pair.savedDigests())[1]);
    } finally {
      pair.destroy();
    }
  }, 120_000);
