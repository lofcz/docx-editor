import { expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  readOoxmlPackage,
  writeOoxmlPackage,
} from '@docx-editor.dev/core/store';
import pairs from '../__tests__/fixtures/ai-pairs.json';
import { C01_SOURCES } from '../../test/support/c01-sources';
import { protectedTargetInput, targetParagraph as p } from '../../test/support/target-document';
import { applyTargetPlanOracle } from '../../test/support/target-plan-oracle';
import { projectSourcePackage } from '../../test/support/target-source-fixture';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { planTargetDocument } from '../diff/plan-target';
import type { CanonicalEdit, CanonicalEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import { readWorkingView } from '../review/working-view';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { applyDirectPlan } from './apply-direct';

const context = { documentId: 'direct-doc', epoch: 'mount-1' };
const token = (editor: DocxEditorInstance) =>
  readEditorRevision(editor, context.documentId, context.epoch)!;
async function projection(editor: DocxEditorInstance) {
  const view = await readWorkingView(editor.surface!.session, token(editor));
  if (!view.success) throw new Error(view.error);
  return view.view.projection;
}
async function plan(
  editor: DocxEditorInstance,
  target: string,
  options: Parameters<typeof planTargetDocument>[2] = {}
): Promise<CanonicalEditPlan> {
  const result = planTargetDocument(
    {
      revision: token(editor),
      pkg: editor.surface!.session.currentPackage(),
      projection: await projection(editor),
    },
    target,
    options
  );
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Semantic oracle: the captured target plans to no operations against the live document. */
async function landedAt(
  editor: DocxEditorInstance,
  target: string,
  options: Parameters<typeof planTargetDocument>[2] = {}
) {
  return (await plan(editor, target, options)).operations;
}
const mount = (document: Uint8Array) =>
  createDocxEditor({
    container: document_(),
    document,
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
const document_ = () => {
  const node = document.createElement('div');
  document.body.append(node);
  return node;
};
const fingerprint = (editor: DocxEditorInstance) =>
  canonicalOoxmlFingerprint(editor.surface!.session.part());

for (const fixture of pairs.cases)
  test(`C01 ${fixture.name}: one direct transaction, one undo step, no review items, exact reopen`, async () => {
    const [body, options] = C01_SOURCES[fixture.name]!(),
      input = await protectedTargetInput(body, options),
      bytes = writeOoxmlPackage(input.pkg),
      editor = mount(bytes);
    try {
      const session = editor.surface!.session,
        before = fingerprint(editor),
        baseline = unzipSync(new Uint8Array(await editor.save())),
        planned = await plan(editor, fixture.afterMarkdown);
      expect(planned.operations.length).toBeGreaterThan(0);
      expect(session.canUndo()).toBe(false);
      let publications = 0;
      const stop = session.subscribe(() => publications++),
        start = token(editor);
      const result = applyDirectPlan(editor, planned, context);
      stop();
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(result.committed).toBe(true);
      expect(publications).toBe(1);
      expect(result.revision.revision).toBe(start.revision + 1);
      expect(editor.getReviewItems()).toEqual([]);
      expect(await landedAt(editor, fixture.afterMarkdown)).toEqual([]);
      // The live executor and the immutable acceptance oracle reach the same native Markdown.
      const oracle = applyTargetPlanOracle(input.pkg, planned);
      const after = (await projection(editor)).text;
      expect(after).toBe((await projectSourcePackage(oracle)).text.replaceAll(/ \/>/g, ' />'));
      const landed = fingerprint(editor);

      expect(session.canUndo()).toBe(true);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(before);
      expect(session.canUndo()).toBe(false);
      expect(await landedAt(editor, fixture.beforeMarkdown)).toEqual([]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(landed);
      expect(session.canRedo()).toBe(false);
      expect(editor.getReviewItems()).toEqual([]);

      const saved = new Uint8Array(await editor.save()),
        original = unzipSync(bytes),
        written = unzipSync(saved);
      // Untouched parts save exactly as an unedited session saves them; media keeps its bytes.
      expect(written['word/media/image.png']).toEqual(original['word/media/image.png']!);
      for (const name of [
        'word/header1.xml',
        'word/footer1.xml',
        'word/footnotes.xml',
        'word/media/image.png',
        'custom/opaque.bin',
      ])
        expect(written[name]).toEqual(baseline[name]!);
      expect(strFromU8(written['word/document.xml']!)).not.toContain('w:ins ');
      const reopened = mount(saved);
      try {
        expect(await landedAt(reopened, fixture.afterMarkdown)).toEqual([]);
        expect((await projection(reopened)).text).toBe(after);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  }, 60_000);

test('stale plans, view and suggestion modes, and stale old values refuse without writes', async () => {
  const input = await protectedTargetInput(p('Keep.') + p('Remove.') + p('Tail.')),
    editor = mount(writeOoxmlPackage(input.pkg));
  try {
    const session = editor.surface!.session,
      before = fingerprint(editor),
      planned = await plan(editor, 'Keep.\n\nChanged.\n\nTail.');
    expect(applyDirectPlan(editor, planned, { ...context, epoch: 'other' })).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    expect(applyDirectPlan(editor, planned, { documentId: '', epoch: 'x' })).toMatchObject({
      success: false,
      code: 'invalid-input',
    });
    for (const [mode, native] of [
      ['suggesting', 'suggest'],
      ['viewing', 'view'],
    ] as const) {
      expect(editor.setEditingMode(mode)).toMatchObject({ ok: true });
      expect(applyDirectPlan(editor, planned, context)).toMatchObject({
        success: false,
        code: 'unsupported-operation',
      });
      expect(editor.surface!.editingMode()).toBe(native);
    }
    expect(editor.setEditingMode('editing').ok).toBe(true);
    // A forged old value refuses before any live write, even with a current revision token.
    const forged: CanonicalEditPlan = {
      ...planned,
      operations: planned.operations.map((edit) =>
        edit.kind === 'text' && edit.before
          ? { ...edit, before: edit.before.replace(/^./u, (c) => (c === 'X' ? 'Y' : 'X')) }
          : edit
      ),
    };
    expect(forged.operations.some((edit, i) => edit !== planned.operations[i])).toBe(true);
    expect(applyDirectPlan(editor, forged, context)).toMatchObject({
      success: false,
      code: 'stale-revision',
    });
    // A failing identity-dependent completion is caught by the full preview.
    const broken: CanonicalEditPlan = {
      ...planned,
      operations: [
        ...planned.operations,
        {
          kind: 'siblings',
          partName: planned.operations[0]!.partName,
          parentId: 'missing-parent',
          beforeId: null,
          afterId: null,
          removedIds: [],
          inserted: [{ type: 'paragraph', children: [{ type: 'text', value: 'X' }] }],
        },
      ],
    };
    expect(applyDirectPlan(editor, broken, context).success).toBe(false);
    expect(fingerprint(editor)).toBe(before);
    expect(session.canUndo()).toBe(false);
    expect(editor.getReviewItems()).toEqual([]);
  } finally {
    editor.destroy();
  }
});

test('an empty plan commits nothing and keeps the revision', async () => {
  const input = await protectedTargetInput(p('Same.')),
    editor = mount(writeOoxmlPackage(input.pkg));
  try {
    const planned = await plan(editor, 'Same.'),
      start = token(editor);
    expect(planned.operations).toEqual([]);
    expect(applyDirectPlan(editor, planned, context)).toEqual({
      success: true,
      revision: start,
      committed: false,
    });
    expect(editor.surface!.session.canUndo()).toBe(false);
  } finally {
    editor.destroy();
  }
});

for (const name of ['complete rewrite', 'tables and links', 'nested MDX and attributes'])
  for (const first of [0, 1] as const)
    test(`two sessions converge on direct C01 ${name} with concurrent peer text; receiver ${first}`, async () => {
      const fixture = pairs.cases.find((item) => item.name === name)!,
        [body, options] = C01_SOURCES[fixture.name]!(),
        input = await protectedTargetInput(body + p('Peer.'), options),
        pair = await createTwoEditorFixture({ document: writeOoxmlPackage(input.pkg) });
      try {
        const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
          target = `${fixture.afterMarkdown}\n\nPeer.`,
          planned = await plan(alice, target),
          peerId = bob.surface!.session.paragraphIds().at(-1)!;
        pair.pause();
        const result = applyDirectPlan(alice, planned, context);
        if (!result.success) throw new Error(result.error);
        expect(
          bob.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: peerId, offset: 5, text: ' Bob' },
          ]).committed
        ).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        const expected = `${fixture.afterMarkdown}\n\nPeer. Bob`;
        for (const editor of pair.editors) {
          expect(await landedAt(editor, expected)).toEqual([]);
          expect(editor.getReviewItems()).toEqual([]);
        }
        // Remote caret geometry follows the peer paragraph after the structural rewrite.
        bob.surface!.setSelection({
          anchor: { paragraphId: peerId, offset: 9 },
          head: { paragraphId: peerId, offset: 9 },
        });
        await pair.settle();
        const remote = pair.rooms[0]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'bob'
        )!;
        expect(remote.head.nodeId).toBe(peerId);
        expect(remote.head.offset).toBe(9);
        const caret = caretAt(alice.surface!.layout(), { paragraphId: peerId, offset: 9 }),
          start = caretAt(alice.surface!.layout(), { paragraphId: peerId, offset: 0 });
        expect(caret !== null && start !== null).toBe(true);
        expect(caret!.x).toBeGreaterThan(start!.x);
        expect(pair.containers[0]!.querySelectorAll('.docx-remote-caret').length).toBeGreaterThan(
          0
        );

        // Undo removes only Alice's transaction and keeps Bob's concurrent text.
        expect(alice.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors)
          expect(await landedAt(editor, `${fixture.beforeMarkdown}\n\nPeer. Bob`)).toEqual([]);
        expect(alice.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) expect(await landedAt(editor, expected)).toEqual([]);

        // Deleting the peer paragraph while Alice redoes elsewhere still converges.
        pair.pause();
        expect(
          bob.surface!.applyAutomationOps(() => [{ op: 'deleteBlock', blockId: peerId }]).committed
        ).toBe(true);
        const firstId = alice.surface!.session.paragraphIds()[0]!;
        expect(
          alice.surface!.applyAutomationOps(() => [
            { op: 'insertText', paragraphId: firstId, offset: 0, text: 'A ' },
          ]).committed
        ).toBe(true);
        pair.reconnect(first);
        await pair.settle();
        const texts = pair.editors.map((editor) =>
          editor
            .surface!.session.paragraphIds()
            .map((id) => paragraphTextOf(editor.surface!.session.part(), id))
        );
        expect(texts[0]).toEqual(texts[1]);
        expect(texts[0]).not.toContain('Peer. Bob');
        const saved = await pair.savedDigests();
        expect(saved[0]).toEqual(saved[1]);
        const reopened = readOoxmlPackage(new Uint8Array(await alice.save()));
        expect(reopened.ok).toBe(true);
      } finally {
        pair.destroy();
      }
    }, 60_000);

const headingStyle = (depth: number) => `Heading${depth}`;
const kinds: readonly [string, string, string, string, readonly CanonicalEdit['kind'][]][] = [
  ['run marks', p('Bold and italic.'), 'Bold *and* italic.', 'Bold and italic.', ['format']],
  ['heading style', p('Title'), '## Title', 'Title', ['format']],
  ['new bullet list', p('One') + p('Two'), '- One\n- Two', 'One\n\nTwo', ['list', 'list']],
  [
    'checklist creation and state',
    p('Task') + p('Done'),
    '- [ ] Task\n- [x] Done',
    'Task\n\nDone',
    ['list', 'list'],
  ],
  [
    'inserted ordered list with a new instance',
    p('Intro.'),
    'Intro.\n\n3. Three\n4. Four',
    'Intro.',
    ['siblings'],
  ],
  [
    'inserted linked paragraph',
    p('Intro.'),
    'Intro.\n\nSee [docs](https://example.org/docs "Docs").',
    'Intro.',
    ['siblings'],
  ],
];
for (const [name, body, target, original, expected] of kinds)
  test(`direct ${name} applies in one undoable transaction and reopens exactly`, async () => {
    const input = await protectedTargetInput(body),
      editor = mount(writeOoxmlPackage(input.pkg));
    try {
      const before = fingerprint(editor),
        planned = await plan(editor, target, { headingStyle });
      expect(planned.operations.map((edit) => edit.kind)).toEqual([...expected]);
      const result = applyDirectPlan(editor, planned, context);
      if (!result.success) throw new Error(`${result.code}: ${result.error}`);
      expect(await landedAt(editor, target, { headingStyle })).toEqual([]);
      expect(editor.getReviewItems()).toEqual([]);
      const landed = fingerprint(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(before);
      expect(editor.surface!.session.canUndo()).toBe(false);
      expect(await landedAt(editor, original, { headingStyle })).toEqual([]);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect(fingerprint(editor)).toBe(landed);
      const reopened = mount(new Uint8Array(await editor.save()));
      try {
        expect(await landedAt(reopened, target, { headingStyle })).toEqual([]);
      } finally {
        reopened.destroy();
      }
    } finally {
      editor.destroy();
    }
  }, 60_000);
