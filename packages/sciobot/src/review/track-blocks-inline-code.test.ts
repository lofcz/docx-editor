import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt, selectionRects } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  collectRevisionSites,
  findNode,
  stylesPartOf,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { docxFixture } from '../../test/support/document';
import {
  blockAst,
  blockContext,
  blockP,
  blockToken,
  blockViews,
  decideBlocks,
  targetBlockPlan,
} from '../../test/support/block-execution';
import { applyInlineCode } from '../native/inline-code';
import { prepareNativeDocument } from '../native/prepare-document';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from './native-module';
import { trackSiblingPlan } from './track-blocks';

const BEFORE = 'Start\\.\n\nTail\\.';
const INSERTED =
  'Plain ``a`<&😀`` with **strong** and [linked `npm test`](https://example.com "Tip")\\.';
const TARGET = `Start\\.\n\n${INSERTED}\n\nTail\\.`;
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const original = docxFixture(blockP('Start.') + blockP('Tail.'), {
  styles: `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Imported"><w:name w:val="Imported"/><w:rPr><w:i/><w:customKeep w:val="style"/></w:rPr></w:style></w:styles>`,
});
const initialized = prepareNativeDocument(original);
if (!initialized.success) throw new Error(initialized.error);
const DOCUMENT = initialized.document;
const moduleOptions = { modules: [SCIOBOT_NATIVE_REVIEW_MODULE] };
const styleFingerprint = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(stylesPartOf(editor.surface!.session.currentPackage())!.root);
async function assertViews(editor: DocxEditorInstance, committed: string, working = committed) {
  expect((await blockViews(editor)).map(blockAst)).toEqual([
    blockAst(committed),
    blockAst(working),
  ]);
}
async function proposal(editor: DocxEditorInstance) {
  const plan = await targetBlockPlan(editor, TARGET);
  expect(plan.operations.every((edit) => edit.kind === 'siblings')).toBe(true);
  const result = trackSiblingPlan(editor, plan, blockContext);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  expect(result.proposals).toHaveLength(1);
  expect(result.proposals[0]!.addresses).toHaveLength(1);
  expect(
    collectRevisionSites(editor.surface!.session.part()).every(
      (site) => site.node.localName === 'ins'
    )
  ).toBe(true);
  return result.proposals;
}
function mount(bytes: Uint8Array) {
  return createDocxEditor({
    container: document.createElement('div'),
    document: bytes,
    ...moduleOptions,
  });
}

for (const decision of ['acceptRevision', 'rejectRevision'] as const)
  test(`tracked new code, emphasis, and links ${decision} exactly through history and reopen`, async () => {
    const editor = mount(DOCUMENT);
    try {
      const part = editor.surface!.session.part();
      const ids = editor.surface!.session.paragraphIds();
      const styleBefore = styleFingerprint(editor);
      const proposals = await proposal(editor);
      await assertViews(editor, BEFORE, TARGET);
      for (const id of ids)
        expect(findNode(editor.surface!.session.part(), id)).toBe(findNode(part, id));
      expect(styleFingerprint(editor)).toBe(styleBefore);
      const pending = new Uint8Array(await editor.save());
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await assertViews(editor, BEFORE);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      await assertViews(editor, BEFORE, TARGET);
      decideBlocks(editor, proposals, decision);
      const resolved = decision === 'acceptRevision' ? TARGET : BEFORE;
      await assertViews(editor, resolved);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      await assertViews(editor, BEFORE, TARGET);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      await assertViews(editor, resolved);
      const reopened = mount(pending);
      const saved = mount(new Uint8Array(await editor.save()));
      try {
        await assertViews(reopened, BEFORE, TARGET);
        decideBlocks(reopened, proposals, decision);
        await assertViews(reopened, resolved);
        await assertViews(saved, resolved);
        expect(styleFingerprint(reopened)).toBe(styleBefore);
        expect(styleFingerprint(saved)).toBe(styleBefore);
      } finally {
        reopened.destroy();
        saved.destroy();
      }
    } finally {
      editor.destroy();
    }
  });

for (const clientIds of [
  [73, 41],
  [41, 73],
] as const)
  for (const receiver of [0, 1] as const)
    for (const decision of ['acceptRevision', 'rejectRevision'] as const)
      test(`tracked new code keeps peer text, geometry, and decisions; clients ${clientIds.join('/')}, receiver ${receiver}, ${decision}`, async () => {
        const pair = await createTwoEditorFixture({
          document: DOCUMENT,
          clientIds,
          ownedAttributes: true,
          editor: moduleOptions,
        });
        try {
          const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance];
          const [first] = alice.surface!.session.paragraphIds();
          const styleBefore = styleFingerprint(alice);
          pair.pause();
          const proposals = await proposal(alice);
          const peer = bob.surface!.applyAutomationOps(
            () => [{ op: 'insertText', paragraphId: first!, offset: 0, text: 'Peer: ' }],
            { kind: 'body' },
            undefined,
            false,
            true
          );
          expect(peer.rejected).toBe(false);
          pair.reconnect(receiver);
          await pair.settle();
          const committed = `Peer: ${BEFORE}`;
          const proposed = `Peer: ${TARGET}`;
          for (const editor of pair.editors) await assertViews(editor, committed, proposed);
          expect(alice.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) await assertViews(editor, committed);
          expect(alice.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) await assertViews(editor, committed, proposed);
          const inserted = alice.surface!.session.paragraphIds()[1]!;
          expect(findNode(bob.surface!.session.part(), inserted)?.kind).toBe('paragraph');
          bob.surface!.setSelection({
            anchor: { paragraphId: inserted, offset: 6 },
            head: { paragraphId: inserted, offset: 12 },
          });
          await pair.settle();
          const remote = pair.rooms[0]!.session.remoteSelections().find(
            (selection) => selection.actorId === 'bob'
          )!;
          expect(remote.anchor.nodeId).toBe(inserted);
          expect(
            caretAt(alice.surface!.layout(), { paragraphId: inserted, offset: remote.head.offset })
          ).toEqual(caretAt(bob.surface!.layout(), { paragraphId: inserted, offset: 12 }));
          const range = {
            anchor: { paragraphId: inserted, offset: 6 },
            head: { paragraphId: inserted, offset: 12 },
          };
          const rectangles = selectionRects(
            alice.surface!.layout(),
            range,
            alice.surface!.session.paragraphIds()
          );
          expect(rectangles.length).toBeGreaterThan(0);
          expect(rectangles).toEqual(
            selectionRects(bob.surface!.layout(), range, bob.surface!.session.paragraphIds())
          );
          decideBlocks(bob, proposals, decision);
          await pair.settle();
          const resolved = decision === 'acceptRevision' ? proposed : committed;
          for (const editor of pair.editors) await assertViews(editor, resolved);
          expect(bob.exec({ type: 'undo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) await assertViews(editor, committed, proposed);
          expect(bob.exec({ type: 'redo' }).ok).toBe(true);
          await pair.settle();
          for (const editor of pair.editors) {
            await assertViews(editor, resolved);
            expect(styleFingerprint(editor)).toBe(styleBefore);
            const reopened = mount(new Uint8Array(await editor.save()));
            try {
              await assertViews(reopened, resolved);
            } finally {
              reopened.destroy();
            }
          }
          if (decision === 'acceptRevision') {
            // A paragraph deletion races a peer format change in its new editable code region.
            pair.pause();
            const code = applyInlineCode(
              alice,
              { paragraphId: inserted, start: 0, end: 5, code: true },
              { ...blockContext, revision: blockToken(alice) }
            );
            expect(code.success).toBe(true);
            expect(
              bob.surface!.applyAutomationOps(
                () => [{ op: 'deleteBlock', blockId: inserted }],
                { kind: 'body' },
                undefined,
                false,
                true
              ).rejected
            ).toBe(false);
            pair.reconnect(receiver);
            await pair.settle();
            for (const editor of pair.editors) {
              await assertViews(editor, committed);
              expect(findNode(editor.surface!.session.part(), inserted)).toBeNull();
            }
            expect(bob.exec({ type: 'undo' }).ok).toBe(true);
            await pair.settle();
            for (const editor of pair.editors)
              await assertViews(editor, proposed.replace('Plain', '`Plain`'));
          }
          const digests = await pair.savedDigests();
          expect(digests[0]).toEqual(digests[1]);
          for (const room of pair.rooms) expect(room.session.status()).toBe('ready');
          pair.ownership.forEach((handle) => handle.assertHealthy());
        } finally {
          pair.destroy();
        }
      }, 60_000);

test('a conflicting imported code style refuses the complete structural proposal without publishing', async () => {
  const source = docxFixture(blockP('Start.') + blockP('Tail.'), {
    styles: `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="SciobotCode"><w:name w:val="Foreign"/></w:style></w:styles>`,
  });
  const editor = mount(source);
  try {
    const before = canonicalOoxmlFingerprint(editor.surface!.session.part());
    const styles = styleFingerprint(editor);
    const revision = blockToken(editor);
    const plan = await targetBlockPlan(editor, TARGET);
    let publications = 0;
    const stop = editor.surface!.session.subscribe(() => publications++);
    const result = trackSiblingPlan(editor, plan, blockContext);
    stop();
    expect(result.success).toBe(false);
    expect(publications).toBe(0);
    expect(blockToken(editor)).toEqual(revision);
    expect(canonicalOoxmlFingerprint(editor.surface!.session.part())).toBe(before);
    expect(styleFingerprint(editor)).toBe(styles);
    expect(editor.surface!.session.canUndo()).toBe(false);
    await assertViews(editor, BEFORE);
  } finally {
    editor.destroy();
  }
});
