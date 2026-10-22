import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  findNode,
  paragraphTextOf,
  readOoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { blockClipboardDocument } from '../../test/support/react-block-document';
import { buildBlockIndex } from './block-index';
import { insertBlock, type BlockInsertion } from './insert';
import { readPayloadStore } from './payload-opc-codec';
import type { BlockPayload } from './payload-schema';

type Element = Exclude<OoxmlNode, { kind: 'textValue' }>;

function control(part: OoxmlPart, id: string): Element | null {
  const visit = (node: OoxmlNode): Element | null => {
    if (node.kind === 'textValue') return null;
    if (node.kind === 'contentControl' && contentControlPropertiesOf(node).tag === `scio:1:${id}`)
      return node;
    for (const child of node.children) {
      const found = visit(child);
      if (found) return found;
    }
    return null;
  };
  return visit(part.root);
}

function content(node: Element): Element {
  return node.children.find(
    (child) => child.kind !== 'textValue' && child.localName === 'sdtContent'
  ) as Element;
}

const ACTIVITY: BlockPayload = {
  id: 'activity-2',
  type: 'lesson_activity',
  version: 1,
  attributes: { name: 'New activity', duration: 5 },
};

function intoPhaseTwo(editor: DocxEditorInstance, extra: Partial<BlockInsertion> = {}) {
  const phase = control(editor.surface!.session.part(), 'phase-2')!;
  return insertBlock(editor, {
    parentControlId: phase.id,
    beforeId: null,
    payload: ACTIVITY,
    label: 'New activity',
    generatedLabel: 'New activity — 5 min',
    ...extra,
  });
}

function payloadOf(editor: DocxEditorInstance, id = 'activity-2') {
  const store = readPayloadStore(editor.surface!.session.currentPackage());
  return store.status === 'ready' ? store.entries.get(id)?.decoded : undefined;
}

function mount() {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: blockClipboardDocument(),
    author: 'Alice',
  });
  return {
    editor,
    destroy() {
      editor.destroy();
      container.remove();
    },
  };
}

test('an activity inserted into a phase is one undo unit with payload, fallback, and caret', async () => {
  const { editor, destroy } = mount();
  try {
    const part = () => editor.surface!.session.part();
    const before = canonicalOoxmlFingerprint(part());
    const beforePackage = editor.surface!.session.currentPackage();
    const result = intoPhaseTwo(editor);
    if (!result.success) throw new Error(result.error);

    const inserted = control(part(), 'activity-2')!;
    expect(inserted.id).toBe(result.controlId);
    const phase = control(part(), 'phase-2')!;
    expect(content(phase).children.at(-1)!.id).toBe(inserted.id);
    expect(contentControlPropertiesOf(inserted).alias).toBe('New activity');
    expect(payloadOf(editor)).toEqual({ status: 'supported', value: ACTIVITY });
    // Generated label, then the caret paragraph.
    const [label, body] = content(inserted).children as Element[];
    expect(contentControlPropertiesOf(label!).tag).toBe('scio:label');
    expect(body!.id).toBe(result.caret.paragraphId);
    expect(editor.surface!.state().selection?.head).toEqual(result.caret);

    // The caret is a real native position: typing lands inside the new block.
    editor.surface!.insertPlainText('Plan');
    editor.surface!.flushPendingInput();
    expect(paragraphTextOf(part(), result.caret.paragraphId)).toBe('Plan');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(paragraphTextOf(part(), result.caret.paragraphId)).toBe('');

    // One more undo removes the block, its fallback, and its payload together.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(canonicalOoxmlFingerprint(part())).toBe(before);
    expect(control(part(), 'activity-2')).toBeNull();
    expect(payloadOf(editor)).toBeUndefined();
    expect(readPayloadStore(editor.surface!.session.currentPackage())).toEqual(
      readPayloadStore(beforePackage)
    );

    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(control(part(), 'activity-2')).not.toBeNull();
    expect(payloadOf(editor)).toEqual({ status: 'supported', value: ACTIVITY });
    const caret = editor.surface!.state().selection!.head;
    expect(findNode(part(), caret.paragraphId)?.kind).toBe('paragraph');
    expect(caretAt(editor.surface!.layout(), caret)).not.toBeNull();

    const reopened = readOoxmlPackage(new Uint8Array(await editor.save()));
    if (!reopened.ok) throw new Error(reopened.reason);
    const index = buildBlockIndex(reopened.package);
    expect(index.diagnostics).toEqual([]);
    const block = index.byId.get('activity-2')!;
    expect(block.parentId).toBe('phase-2');
    expect(block.payload).toEqual(ACTIVITY);
  } finally {
    destroy();
  }
});

test('body insertion lands before a sibling or at the end, before section properties', () => {
  const { editor, destroy } = mount();
  try {
    const part = () => editor.surface!.session.part();
    const phase = control(part(), 'phase-2')!;
    const first = insertBlock(editor, {
      parentControlId: null,
      beforeId: phase.id,
      payload: { ...ACTIVITY, id: 'before-phase' },
      label: 'Before',
      paragraphs: ['First line', 'Second line'],
    });
    if (!first.success) throw new Error(first.error);
    const last = insertBlock(editor, {
      parentControlId: null,
      beforeId: null,
      payload: { ...ACTIVITY, id: 'at-end' },
      label: 'End',
    });
    if (!last.success) throw new Error(last.error);
    const body = part().root.children[0] as Element;
    const ids = body.children.map((child) => child.id);
    expect(ids.indexOf(first.controlId)).toBe(ids.indexOf(phase.id) - 1);
    const terminal = body.children.at(-1)!;
    const hasSection = terminal.kind !== 'textValue' && terminal.localName === 'sectPr';
    expect(ids.at(hasSection ? -2 : -1)).toBe(last.controlId);
    expect(paragraphTextOf(part(), first.caret.paragraphId)).toBe('First line');
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(control(part(), 'at-end')).toBeNull();
    expect(control(part(), 'before-phase')).not.toBeNull();
  } finally {
    destroy();
  }
});

test('invalid insertions refuse before any native write', () => {
  const { editor, destroy } = mount();
  try {
    const part = () => editor.surface!.session.part();
    const before = canonicalOoxmlFingerprint(part());
    const phase = control(part(), 'phase-2')!;
    const activity = control(part(), 'activity-1')!;
    const cases: [Partial<BlockInsertion>, string][] = [
      [{ payload: { ...ACTIVITY, id: 'activity-1' } }, 'invalid-input'],
      [{ payload: { ...ACTIVITY, id: 'bad id' } }, 'invalid-input'],
      [{ parentControlId: 'missing' }, 'invalid-input'],
      [{ beforeId: activity.id }, 'invalid-input'],
      [{ paragraphs: [] }, 'invalid-input'],
      [{ paragraphs: ['two\nlines'] }, 'invalid-input'],
      [{ expectedRevision: -1 }, 'stale-revision'],
    ];
    for (const [extra, code] of cases) {
      const result = intoPhaseTwo(editor, { parentControlId: phase.id, ...extra });
      expect(result).toMatchObject({ success: false, code });
      expect(canonicalOoxmlFingerprint(part())).toBe(before);
    }
    editor.surface!.setEditingMode('suggest');
    expect(intoPhaseTwo(editor)).toMatchObject({ success: false, code: 'unsupported-operation' });
    editor.surface!.setEditingMode('view');
    expect(intoPhaseTwo(editor).success).toBe(false);
    expect(canonicalOoxmlFingerprint(part())).toBe(before);
    expect(payloadOf(editor)).toBeUndefined();
  } finally {
    destroy();
  }
});

for (const clientIds of [
  [73, 41],
  [41, 73],
] as const)
  for (const first of [0, 1] as const)
    test(`two sessions converge on an inserted activity beside offline peer typing; clients ${clientIds.join('/')}, receiver ${first}`, async () => {
      const pair = await createTwoEditorFixture({
        document: blockClipboardDocument(),
        clientIds,
      });
      try {
        const [alice, bob] = pair.editors;
        const bobPhase = control(bob!.surface!.session.part(), 'phase-2')!;
        const phaseText = (content(bobPhase).children[0] as Element).id;
        pair.pause();
        const result = intoPhaseTwo(alice!);
        if (!result.success) throw new Error(result.error);
        bob!.surface!.setSelection({
          anchor: { paragraphId: phaseText, offset: 0 },
          head: { paragraphId: phaseText, offset: 0 },
        });
        bob!.surface!.insertPlainText('Bob: ');
        pair.reconnect(first);
        await pair.settle();
        for (const editor of pair.editors) {
          const part = editor.surface!.session.part();
          expect(control(part, 'activity-2')).not.toBeNull();
          expect(payloadOf(editor)).toEqual({ status: 'supported', value: ACTIVITY });
          expect(paragraphTextOf(part, phaseText)).toBe('Bob: Phase two');
        }
        const [savedAlice, savedBob] = await pair.savedDigests();
        expect(savedAlice).toEqual(savedBob);

        // Alice types in the new block; Bob sees her caret there.
        const caretParagraph = alice!.surface!.state().selection!.head.paragraphId;
        alice!.surface!.insertPlainText('Alice');
        await pair.settle();
        const remote = pair.rooms[1]!.session.remoteSelections().find(
          (selection) => selection.actorId === 'alice'
        );
        expect(Boolean(remote)).toBe(true);
        expect(
          caretAt(bob!.surface!.layout(), {
            paragraphId: remote!.head.nodeId,
            offset: remote!.head.offset,
          })
        ).not.toBeNull();
        for (const editor of pair.editors)
          expect(
            paragraphTextOf(
              editor.surface!.session.part(),
              (
                content(control(editor.surface!.session.part(), 'activity-2')!)
                  .children[1] as Element
              ).id
            )
          ).toBe('Alice');
        expect(paragraphTextOf(alice!.surface!.session.part(), caretParagraph)).toBe('Alice');

        // Undo Alice's typing, then her insertion; Bob's text survives. Redo restores both.
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        expect(alice!.exec({ type: 'undo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(control(editor.surface!.session.part(), 'activity-2')).toBeNull();
          expect(payloadOf(editor)).toBeUndefined();
          expect(paragraphTextOf(editor.surface!.session.part(), phaseText)).toBe('Bob: Phase two');
        }
        expect(alice!.exec({ type: 'redo' }).ok).toBe(true);
        await pair.settle();
        for (const editor of pair.editors) {
          expect(control(editor.surface!.session.part(), 'activity-2')).not.toBeNull();
          expect(payloadOf(editor)).toEqual({ status: 'supported', value: ACTIVITY });
        }
        const [redoAlice, redoBob] = await pair.savedDigests();
        expect(redoAlice).toEqual(redoBob);
      } finally {
        pair.destroy();
      }
    }, 60_000);

test('a peer deleting the parent phase during insertion leaves both sessions equal', async () => {
  const pair = await createTwoEditorFixture({
    document: blockClipboardDocument(),
    clientIds: [41, 73],
  });
  try {
    const [alice, bob] = pair.editors;
    pair.pause();
    expect(intoPhaseTwo(alice!).success).toBe(true);
    const phase = control(bob!.surface!.session.part(), 'phase-2')!;
    expect(
      bob!.surface!.applyAutomationOps(() => [
        { op: 'removeContentControl', controlId: phase.id, keepContent: false },
      ]).rejected
    ).toBe(false);
    pair.reconnect(0);
    await pair.settle();
    for (const room of pair.rooms) expect(room.session.status()).not.toBe('error');
    const [savedAlice, savedBob] = await pair.savedDigests();
    expect(savedAlice).toEqual(savedBob);
    expect(Boolean(control(alice!.surface!.session.part(), 'activity-2'))).toBe(
      Boolean(control(bob!.surface!.session.part(), 'activity-2'))
    );
  } finally {
    pair.destroy();
  }
}, 60_000);
