import { beginHistoryGroupFrames, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import {
  paragraphTextOf,
  parentNodeOf,
  findNode,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { applyDirectListAssignments } from '../ai/direct-lists';
import { planParagraphText } from '../diff/paragraph-text';
import { planListMembership } from '../diff/lists';
import { createEditPlan } from '../diff/plan';
import { readEditorRevision } from '../projection/revision-cache';
import type { ListTarget } from '../diff/list-types';
import type { BlockCommandContext } from './block-type';
import {
  applyManualListPlan,
  manualListState,
  adjustManualListLevel,
  setManualChecklist,
} from './lists';

/** Consume exactly one captured prefix and apply membership in the same canonical transaction. */
export function applyManualListInput(
  editor: DocxEditorInstance,
  context: BlockCommandContext
): boolean {
  const s = editor.surface;
  if (!s) return false;
  s.flushPendingInput();
  const { anchor, head } = s.state().selection;
  if (anchor.paragraphId !== head.paragraphId || anchor.offset !== head.offset) return false;
  const part = s.session.part(),
    text = paragraphTextOf(part, head.paragraphId),
    revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (text === null || text === undefined || text.length !== head.offset || !revision) return false;
  const current = manualListState(editor, context);
  if (!current || current.numbering || current.checkbox) return false;
  let target: ListTarget;
  const ordered = /^(\d{1,9})[.)]$/.exec(text);
  if (text === '-' || text === '*')
    target = {
      kind: 'bullet',
      style: 'disc',
      level: 0,
      instance: { kind: 'new', key: `input:${revision.revision}` },
    };
  else if (ordered)
    target = {
      kind: 'ordered',
      style: 'decimal',
      level: 0,
      instance: { kind: 'new', key: `input:${revision.revision}`, start: Number(ordered[1]) },
    };
  else if (['[]', '[ ]', '[x]', '[X]'].includes(text))
    target = { kind: 'checklist', level: 0, checked: /x/i.test(text) };
  else return false;
  const list = planListMembership(
    { pkg: s.session.currentPackage(), partName: part.name, revision },
    [{ paragraphId: head.paragraphId, target }]
  );
  const removal = planParagraphText(
    { partName: part.name, paragraphId: head.paragraphId, revision, text },
    ''
  );
  if (!list.success || !removal.success) return false;
  const plan = createEditPlan(revision, [...removal.plan.operations, ...list.plan.operations]);
  return plan.success && applyManualListPlan(editor, plan.plan, context).success;
}

/** Continue checklist metadata through the native paragraph split; Enter at the end resets checked. */
export function handleManualListKey(
  editor: DocxEditorInstance,
  key: 'Enter' | 'Backspace',
  context: BlockCommandContext
): boolean {
  const s = editor.surface;
  if (!s || !editor.snapshot().editable) return false;
  s.flushPendingInput();
  const { anchor, head } = s.state().selection;
  if (anchor.paragraphId !== head.paragraphId) return false;
  const state = manualListState(editor, context),
    text = paragraphTextOf(s.session.part(), head.paragraphId);
  if (!state || text === null || (!state.checkbox && !state.numbering)) return false;
  const collapsed = anchor.offset === head.offset,
    marker = state.checkbox ? 1 : 0;
  if (key === 'Backspace')
    return (
      collapsed &&
      head.offset <= marker &&
      adjustManualListLevel(editor, 'decrease', context).success
    );
  if (collapsed && text!.length === marker)
    return adjustManualListLevel(editor, 'decrease', context).success;
  if (!state.checkbox) return false;
  if (s.editingMode() === 'edit') {
    const source = s.session.part(),
      parent = parentNodeOf(source, head.paragraphId);
    if (!parent) return false;
    const start = Math.min(anchor.offset, head.offset),
      end = Math.max(anchor.offset, head.offset);
    const checked = collapsed && head.offset === text!.length ? false : state.checkbox.checked;
    const next = (part: typeof source) => {
      const container = findNode(part, parent.id);
      if (!container || container.kind === 'textValue')
        throw new Error('The checklist paragraph parent disappeared.');
      const at = container.children.findIndex((node) => node.id === head.paragraphId);
      const paragraph = container.children[at + 1];
      if (paragraph?.kind !== 'paragraph')
        throw new Error('The native checklist split did not allocate a paragraph.');
      return paragraph.id;
    };
    let committed = false;
    s.commitReviewOps(
      () => {
        const ops: TreeDocOp[] = [
          ...(!collapsed
            ? [{ op: 'deleteText' as const, paragraphId: head.paragraphId, start, end }]
            : []),
          { op: 'splitParagraph', paragraphId: head.paragraphId, offset: start },
        ];
        const result = s.applyAutomationOps(
          () => ops,
          { kind: 'body' },
          [
            (pkg) =>
              applyDirectListAssignments(pkg, source.name, [
                {
                  paragraphId: next(pkg.parts.get(source.name)!),
                  before: null,
                  target: { kind: 'checklist', checked, level: state.checkbox!.level },
                },
              ]),
          ],
          false,
          true,
          'after'
        );
        committed = result.committed && !result.rejected;
        return result;
      },
      'revision-resolve',
      () => {
        if (!committed) return null;
        const point = { paragraphId: next(s.session.part()), offset: 1 };
        return { anchor: point, head: point };
      }
    );
    return committed;
  }
  const group = editor.beginHistoryGroup(),
    frames = beginHistoryGroupFrames(editor, group);
  if (!frames) {
    group.end();
    return false;
  }
  try {
    frames.run(() => s.splitParagraph());
    if (editor.surface !== s || s.state().selection.head.paragraphId === head.paragraphId)
      return true;
    const checked = collapsed && head.offset === text!.length ? false : state.checkbox.checked;
    frames.run(() => setManualChecklist(editor, checked, state.checkbox!.level, context));
    return true;
  } finally {
    frames.finish();
    group.end();
  }
}

export function installListEditing(editor: DocxEditorInstance, context: BlockCommandContext) {
  const s = editor.surface;
  if (!s) throw new Error('The native editor is not attached.');
  const host = s.hostElement();
  const inSurface = (event: Event) => {
    const target = event.target as Element | null;
    return (
      editor.surface === s &&
      !!target?.closest?.('[contenteditable="true"]') &&
      !target.closest('[data-docx-block-ui]')
    );
  };
  const stop = (event: Event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
  };
  const input = (event: InputEvent) => {
    if (
      inSurface(event) &&
      !event.isComposing &&
      event.inputType === 'insertText' &&
      event.data === ' ' &&
      applyManualListInput(editor, context)
    )
      stop(event);
  };
  const key = (event: KeyboardEvent) => {
    if (!inSurface(event) || event.isComposing || event.altKey || event.ctrlKey || event.metaKey)
      return;
    if (event.key === 'Tab') {
      const state = manualListState(editor, context);
      if (
        state?.checkbox &&
        adjustManualListLevel(editor, event.shiftKey ? 'decrease' : 'increase', context).success
      )
        stop(event);
    } else if (
      !event.shiftKey &&
      (event.key === 'Enter' || event.key === 'Backspace') &&
      handleManualListKey(editor, event.key, context)
    )
      stop(event);
  };
  host.addEventListener('beforeinput', input, true);
  host.addEventListener('keydown', key, true);
  return () => {
    host.removeEventListener('beforeinput', input, true);
    host.removeEventListener('keydown', key, true);
  };
}
