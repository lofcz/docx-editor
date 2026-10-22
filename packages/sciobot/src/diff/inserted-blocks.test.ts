import { expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  contentControlPropertiesOf,
  contentControlsIn,
  paragraphTextOf,
  writeOoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { C01_SOURCES } from '../../test/support/c01-sources';
import { protectedTargetInput } from '../../test/support/target-document';
import { createTwoEditorFixture } from '../../test/support/collaboration';
import { buildBlockIndex } from '../blocks/block-index';
import { blockTag } from '../blocks/payload-opc-codec';
import { readEditorRevision } from '../projection/revision-cache';
import { SCIOBOT_NATIVE_REVIEW_MODULE } from '../review/native-module';
import { readWorkingView } from '../review/working-view';
import { applyDirectPlan } from '../ai/apply-direct';
import { decodeEditPlan, serializeEditPlan, type CanonicalEditPlan } from './plan';
import { planTargetDocument, type TargetDocumentOptions } from './plan-target';

const context = { documentId: 'blocks', epoch: 'mount' };
const ACTIVITY =
  '<activity name="Warmup" duration="10">\nFirst question.\n\n<source title="Source" url="https://example.com" type="verified" />\n\nLast question.\n</activity>';
const lesson = (phase: string, after = '') =>
  `<lesson_title grade="7">Lesson</lesson_title>\n\n<phase type="evocation">\n${phase}\n</phase>${after}`;
let counter = 0;
const ids = (): TargetDocumentOptions => {
  const prefix = `new${++counter}`;
  let n = 0;
  return { allocateBlockId: () => `${prefix}-${++n}` };
};
async function mount() {
  const [body, options] = C01_SOURCES['nested MDX and attributes']!(),
    container = document.createElement('div');
  document.body.append(container);
  return createDocxEditor({
    container,
    document: writeOoxmlPackage((await protectedTargetInput(body, options)).pkg),
    author: 'Teacher',
    modules: [SCIOBOT_NATIVE_REVIEW_MODULE],
  });
}
async function source(editor: DocxEditorInstance) {
  const token = readEditorRevision(editor, context.documentId, context.epoch)!,
    view = await readWorkingView(editor.surface!.session, token);
  if (!view.success) throw new Error(view.error);
  return {
    revision: token,
    pkg: editor.surface!.session.currentPackage(),
    projection: view.view.projection,
  };
}
async function plan(
  editor: DocxEditorInstance,
  target: string,
  options: TargetDocumentOptions = {}
): Promise<CanonicalEditPlan> {
  const result = planTargetDocument(await source(editor), target, options);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  return result.plan;
}
/** Direct landing, then empty replanning, one undo step, redo, and reopen as semantic oracles. */
async function land(editor: DocxEditorInstance, target: string, planned: CanonicalEditPlan) {
  expect(decodeEditPlan(serializeEditPlan(planned))).toEqual({ success: true, plan: planned });
  const result = applyDirectPlan(editor, planned, context);
  if (!result.success) throw new Error(`${result.code}: ${result.error}`);
  expect((await plan(editor, target)).operations).toEqual([]);
  const landed = buildBlockIndex(editor.surface!.session.currentPackage());
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.surface!.session.canUndo()).toBe(false);
  expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(4);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  const container = document.createElement('div');
  document.body.append(container);
  const reopened = createDocxEditor({ container, document: new Uint8Array(await editor.save()) });
  try {
    expect((await plan(reopened, target)).operations).toEqual([]);
    expect(buildBlockIndex(reopened.surface!.session.currentPackage()).byId.size).toBe(
      landed.byId.size
    );
  } finally {
    reopened.destroy();
    container.remove();
  }
  return landed;
}
const textsOf = (editor: DocxEditorInstance, controlId: string) => {
  const part = editor.surface!.session.part();
  const control = contentControlsIn(part.root).find((entry) => entry.node.id === controlId)!;
  const out: string[] = [];
  const visit = (node: OoxmlNode) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') out.push(paragraphTextOf(part, node.id)!);
    else node.children.forEach(visit);
  };
  visit(control.node);
  return out;
};

test('a new activity in an existing phase carries its payload, marks, link, and nested list', async () => {
  const editor = await mount();
  try {
    const target = lesson(
        `${ACTIVITY}\n\n<activity name="Discussion" duration="5">\nTalk about **it** with [notes](https://example.org).\n\n- One\n- Two\n</activity>`
      ),
      options = ids(),
      planned = await plan(editor, target, options);
    const sibling = planned.operations.find((edit) => edit.kind === 'siblings');
    expect(sibling).toMatchObject({
      kind: 'siblings',
      removedIds: [],
      blocks: [
        {
          path: [0],
          label: 'activity',
          payload: {
            version: 1,
            type: 'lesson_activity',
            attributes: { name: 'Discussion', duration: 5 },
          },
        },
      ],
      lists: [
        { path: [0, 1], target: { kind: 'bullet', level: 0 } },
        { path: [0, 2], target: { kind: 'bullet', level: 0 } },
      ],
    });
    const id = sibling!.kind === 'siblings' ? sibling!.blocks![0]!.payload.id : '';
    const before = buildBlockIndex(editor.surface!.session.currentPackage());
    const landed = await land(editor, target, planned);
    const block = landed.byId.get(id)!;
    expect(block.payload).toEqual({
      version: 1,
      id,
      type: 'lesson_activity',
      attributes: { name: 'Discussion', duration: 5 },
    });
    expect(block.parentId).toBe('phase');
    const control = contentControlsIn(editor.surface!.session.part().root).find(
      (entry) => entry.node.id === block.controlId
    )!;
    expect(contentControlPropertiesOf(control.node).tag).toBe(blockTag(id));
    expect(textsOf(editor, block.controlId)).toEqual(['Talk about it with notes.', 'One', 'Two']);
    for (const old of ['title', 'phase', 'activity', 'source'])
      expect(landed.byId.get(old)!.controlId).toBe(before.byId.get(old)!.controlId);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a new phase with a nested activity and a new empty container land in one transaction', async () => {
  const editor = await mount();
  try {
    const target = lesson(
      ACTIVITY,
      '\n\n<phase type="reflection">\n<activity name="Wrap" duration="3">\nWhat did we learn?\n</activity>\n</phase>\n\n<recap_goals>\n</recap_goals>'
    );
    const planned = await plan(editor, target, ids()),
      sibling = planned.operations.find((edit) => edit.kind === 'siblings')!;
    expect(
      sibling.kind === 'siblings' &&
        sibling.blocks!.map((binding) => [binding.path, binding.payload.type])
    ).toEqual([
      [[0], 'lesson_phase'],
      [[0, 0], 'lesson_activity'],
      [[1], 'recap_goals'],
    ]);
    const landed = await land(editor, target, planned),
      [phase, activity, goals] = sibling.kind === 'siblings' ? sibling.blocks! : [];
    expect(landed.byId.get(activity!.payload.id)!.parentId).toBe(phase!.payload.id);
    expect(landed.byId.get(phase!.payload.id)!.payload.attributes).toEqual({
      phaseType: 'reflection',
    });
    expect(textsOf(editor, landed.byId.get(goals!.payload.id)!.controlId)).toEqual(['']);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('a new atomic visual stores its payload with a readable Word fallback', async () => {
  const editor = await mount();
  try {
    const target = lesson(
      `${ACTIVITY}\n\n<lesson_visual kind="image" src="https://example.org/a.png" alt="Diagram" />`
    );
    const planned = await plan(editor, target, ids()),
      sibling = planned.operations.find((edit) => edit.kind === 'siblings')!,
      binding = sibling.kind === 'siblings' ? sibling.blocks![0]! : null;
    expect(binding!.payload).toMatchObject({
      type: 'lesson_visual',
      attributes: { kind: 'image', src: 'https://example.org/a.png', alt: 'Diagram' },
    });
    const landed = await land(editor, target, planned);
    expect(textsOf(editor, landed.byId.get(binding!.payload.id)!.controlId)).toEqual(['Diagram']);
  } finally {
    editor.destroy();
  }
}, 60_000);

test('unknown tags, reused IDs, and invalid allocations refuse without a plan', async () => {
  const editor = await mount();
  try {
    const input = await source(editor),
      activity = `${ACTIVITY}\n\n<activity name="New" duration="1">\nText.\n</activity>`;
    for (const [target, options] of [
      [lesson(`${ACTIVITY}\n\n<Unknown value={danger()} />`), ids()],
      [lesson(activity), { allocateBlockId: () => 'activity' }],
      [lesson(activity), { allocateBlockId: () => 'not valid!' }],
      [lesson(`${ACTIVITY}\n\n<activity name="Q" duration="1">\n> quoted\n</activity>`), ids()],
    ] as const) {
      const result = planTargetDocument(input, target, options);
      expect({ target, success: result.success }).toEqual({ target, success: false });
      expect(Object.hasOwn(result, 'plan')).toBe(false);
    }
    expect((globalThis as Record<string, unknown>).danger).toBeUndefined();
  } finally {
    editor.destroy();
  }
}, 60_000);

for (const first of [0, 1] as const)
  test(`two sessions converge on a new activity beside concurrent peer text; receiver ${first}`, async () => {
    const [body, options] = C01_SOURCES['nested MDX and attributes']!(),
      pair = await createTwoEditorFixture({
        document: writeOoxmlPackage((await protectedTargetInput(body, options)).pkg),
      });
    try {
      const [alice, bob] = pair.editors as [DocxEditorInstance, DocxEditorInstance],
        target = lesson(
          `${ACTIVITY}\n\n<activity name="New" duration="2">\nFresh text.\n</activity>`
        ),
        planned = await plan(alice, target, ids()),
        part = bob.surface!.session.part(),
        last = bob
          .surface!.session.paragraphIds()
          .find((id) => paragraphTextOf(part, id) === 'Last question.')!;
      pair.pause();
      const result = applyDirectPlan(alice, planned, context);
      if (!result.success) throw new Error(result.error);
      expect(
        bob.surface!.applyAutomationOps(() => [
          { op: 'insertText', paragraphId: last, offset: 14, text: ' Bob' },
        ]).committed
      ).toBe(true);
      pair.reconnect(first);
      await pair.settle();
      const expected = target.replace('Last question.', 'Last question. Bob');
      for (const editor of pair.editors) {
        expect((await plan(editor, expected)).operations).toEqual([]);
        expect(buildBlockIndex(editor.surface!.session.currentPackage()).byId.size).toBe(5);
      }
      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      await pair.settle();
      const undone = lesson(ACTIVITY).replace('Last question.', 'Last question. Bob');
      for (const editor of pair.editors)
        expect((await plan(editor, undone)).operations).toEqual([]);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      await pair.settle();
      for (const editor of pair.editors)
        expect((await plan(editor, expected)).operations).toEqual([]);
      const [savedAlice, savedBob] = await pair.savedDigests();
      expect(savedAlice).toEqual(savedBob);
    } finally {
      pair.destroy();
    }
  }, 60_000);
