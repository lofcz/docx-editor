import { listToggleSiblings } from './list-toggle-siblings';
import { paragraphTextOf, paragraphOffsetIndex, findNode } from '@docx-editor.dev/core/store';
import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphsInCells, paragraphFragmentsOf } from '@docx-editor.dev/core/layout';
import { createListSourceReader, planListMembership } from '../diff/lists';
import { planManualListSequences } from './list-sequences';
import { planParagraphFormat } from '../diff/paragraph-format';
import { createEditPlan, type CanonicalEdit } from '../diff/plan';
import { insertManualList } from './insert-list';
import { paragraphIndentPatch } from './paragraph-position';
import {
  LIST_STYLES,
  type ListStyle,
  type ListTarget,
  type ListSourceState,
} from '../diff/list-types';
import { applyDirectPlan } from '../ai/apply-direct';
import { replaceOwnedProposals } from '../review/replace-owned-proposals';
import { readEditorRevision } from '../projection/revision-cache';
import type { BlockCommandContext } from './block-type';
export { LIST_STYLES };
const fail = (error: string) => ({
  success: false as const,
  code: 'unsupported-operation' as const,
  error,
});
function selection(editor: DocxEditorInstance) {
  const s = editor.surface!;
  const { anchor, head } = s.state().selection,
    order = s.session.paragraphIds(),
    a = order.indexOf(anchor.paragraphId),
    h = order.indexOf(head.paragraphId),
    end = a <= h ? head : anchor;
  const cells = editor.getTableCellSelection()?.cellIds;
  return cells?.length
    ? paragraphsInCells(s.layout(), cells)
    : order.slice(Math.min(a, h), Math.max(a, h) + (a !== h && end.offset === 0 ? 0 : 1));
}
/** Read exactly the canonical list source used by the planner, including owned Word checkboxes. */
export function manualListState(
  editor: DocxEditorInstance,
  context: BlockCommandContext,
  paragraphId = editor.surface?.state().selection.head.paragraphId
): ListSourceState | null {
  const s = editor.surface,
    revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (!s || !revision || !paragraphId) return null;
  const reader = createListSourceReader(s.session.currentPackage(), s.session.part());
  const inspected = reader.read(paragraphId);
  return inspected.success ? inspected.state : null;
}

function change(
  editor: DocxEditorInstance,
  context: BlockCommandContext,
  target: (state: ListSourceState, id: string) => ListTarget | ReturnType<typeof fail>,
  sequences = false,
  targetIds?: readonly string[]
) {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  const revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (!revision) return fail('The native document is not ready.');
  const pkg = s.session.currentPackage(),
    part = s.session.partFor(s.storyScope()) ?? s.session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('List controls require the body story.');
  const patches = [];
  const reader = createListSourceReader(pkg, part);
  for (const id of targetIds ?? selection(editor)) {
    const inspected = reader.read(id);
    if (!inspected.success) return inspected;
    const state = inspected.state;
    const next = target(state, id);
    if ('success' in next) return next;
    patches.push({ paragraphId: id, target: next });
  }
  const planned = (sequences ? planManualListSequences : planListMembership)(
    { pkg, partName: part.name, revision },
    patches
  );
  if (!planned.success) return planned;
  if (!planned.plan.operations.length)
    return { success: true as const, committed: false, revision };
  return applyManualListPlan(editor, planned.plan, context, sequences);
}
export function applyManualListPlan(
  editor: DocxEditorInstance,
  plan: Parameters<typeof applyDirectPlan>[1],
  context: BlockCommandContext,
  groupNativeChanges = false
) {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  return s.editingMode() === 'suggest'
    ? replaceOwnedProposals(
        editor,
        [],
        plan,
        {
          ...context,
          author: editor.getConfiguredAuthor() ?? '',
          date: new Date().toISOString(),
        },
        {},
        undefined,
        groupNativeChanges
      )
    : commitDirectMembership(editor, plan, context);
}
/** The native guarded boundary maps the caret across an owned checkbox atom and captures history. */
function commitDirectMembership(
  editor: DocxEditorInstance,
  plan: Parameters<typeof applyDirectPlan>[1],
  context: BlockCommandContext
) {
  let result: ReturnType<typeof applyDirectPlan> | undefined;
  const surface = editor.surface!,
    before = surface.session.part(),
    selected = surface.state().selection;
  const mapped = () => {
    const map = (point: typeof selected.head) => {
      const old = paragraphTextOf(before, point.paragraphId),
        current = paragraphTextOf(surface.session.part(), point.paragraphId);
      if (old === null || current === null) return point;
      const edit = plan.operations.find(
        (e) => e.kind === 'list' && e.paragraphId === point.paragraphId
      );
      if (edit?.kind === 'list' && edit.before.checkbox) {
        const paragraph = findNode(before, point.paragraphId);
        const start =
          paragraph?.kind === 'paragraph'
            ? paragraphOffsetIndex(paragraph).spanOf(edit.before.checkbox.controlId)?.start
            : undefined;
        if (start !== undefined && point.offset <= start) return point;
      }
      return {
        ...point,
        offset: Math.max(0, Math.min(current.length, point.offset + current.length - old.length)),
      };
    };
    return { anchor: map(selected.anchor), head: map(selected.head) };
  };
  const refusal = surface.commitReviewOps(
    () => {
      result = applyDirectPlan(editor, plan, context);
      return {
        committed: result.success && result.committed,
        ...(!result.success ? { reason: result.error } : {}),
      };
    },
    'revision-resolve',
    mapped
  );
  return result ?? fail(refusal ?? 'The native list edit did not commit.');
}
export function setManualList(
  editor: DocxEditorInstance,
  style: ListStyle | 'todo' | 'none',
  context: BlockCommandContext
) {
  return setListTargets(editor, style, context);
}
function setListTargets(
  editor: DocxEditorInstance,
  style: ListStyle | 'todo' | 'none',
  context: BlockCommandContext,
  targetIds?: readonly string[]
) {
  if (![...LIST_STYLES, 'todo', 'none'].includes(style))
    return fail('Choose an application list style.');
  editor.surface?.flushPendingInput();
  const key = `manual:${editor.surface?.session.packageRevision() ?? 0}`;
  return change(
    editor,
    context,
    (state) => {
      if (style === 'none') return { kind: 'none' };
      const level = state.checkbox?.level ?? state.numbering?.level ?? 0;
      return style === 'todo'
        ? { kind: 'checklist', level, checked: state.checkbox?.checked ?? false }
        : {
            kind: ['disc', 'circle', 'square'].includes(style) ? 'bullet' : 'ordered',
            level,
            instance: { kind: 'new', key },
            style,
          };
    },
    true,
    targetIds
  );
}
/** Toggle native membership; captured checklist indentation has a separate dispatch contract. */
export function toggleManualList(
  editor: DocxEditorInstance,
  style: ListStyle | 'todo',
  context: BlockCommandContext
) {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  const formats: Readonly<Record<ListStyle, string>> = {
    disc: 'bullet',
    circle: 'bullet',
    square: 'bullet',
    decimal: 'decimal',
    'lower-alpha': 'lowerLetter',
    'upper-alpha': 'upperLetter',
    'lower-roman': 'lowerRoman',
    'upper-roman': 'upperRoman',
  };
  const glyphs: Readonly<Record<string, string>> = { disc: '•', circle: '○', square: '■' };
  const ids = selection(editor);
  const active =
    ids.length > 0 &&
    ids.every((id) => {
      const state = manualListState(editor, context, id),
        marker = manualListMarker(editor, id);
      return style === 'todo'
        ? !!state?.checkbox
        : !state?.checkbox &&
            marker?.numFmt === formats[style] &&
            (!glyphs[style] || marker.text === glyphs[style]);
    });
  if (!active && selectedCollapsed(editor)) {
    const revision = readEditorRevision(editor, context.documentId, context.epoch);
    if (!revision) return fail('The native document is not ready.');
    try {
      const targetIds = listToggleSiblings(
        { pkg: s.session.currentPackage(), partName: s.session.part().name, revision },
        ids[0]!
      );
      return setListTargets(editor, style, context, targetIds);
    } catch (error) {
      return fail(error instanceof Error ? error.message : String(error));
    }
  }
  if (active) return removeManualListToggle(editor, ids, context);
  return setManualList(editor, style, context);
}
function selectedCollapsed(editor: DocxEditorInstance) {
  const { anchor, head } = editor.surface!.state().selection;
  return anchor.paragraphId === head.paragraphId && anchor.offset === head.offset;
}
function removeManualListToggle(
  editor: DocxEditorInstance,
  ids: readonly string[],
  context: BlockCommandContext
) {
  const s = editor.surface!,
    pkg = s.session.currentPackage(),
    part = pkg.parts.get(pkg.mainDocumentPart)!;
  const revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (!revision) return fail('The native document is not ready.');
  const patches = [];
  for (const paragraphId of ids) {
    const state = manualListState(editor, context, paragraphId);
    if (!state) return fail('The selected list source is unavailable.');
    patches.push({
      paragraphId,
      target: {
        kind: 'none' as const,
        retainedIndent: state.checkbox?.level ?? state.numbering?.level ?? 0,
      },
    });
  }
  const planned = planManualListSequences({ pkg, partName: part.name, revision }, patches);
  if (!planned.success) return planned;
  if (!planned.plan.operations.length)
    return { success: true as const, committed: false, revision };
  return applyManualListPlan(editor, planned.plan, context, true);
}
/** Match the captured indentTodo action, independently of native checkbox conversion. */
export function indentManualChecklist(editor: DocxEditorInstance, context: BlockCommandContext) {
  editor.surface?.flushPendingInput();
  const key = `checklist-indent:${editor.surface?.session.packageRevision() ?? 0}`;
  return change(editor, context, (state, id) => {
    let level = (state.checkbox?.level ?? state.numbering?.level ?? -1) + 1;
    if (!state.checkbox && !state.numbering) {
      const paragraph = editor.surface
        ?.layout()
        .pages.flatMap((page) => paragraphFragmentsOf(page))
        .find((fragment) => fragment.paragraphId === id);
      if (!paragraph) return fail('The selected paragraph has no native layout.');
      level = Math.round(paragraph.indent.left * 20) / 360;
    }
    if (!Number.isInteger(level) || level < 0 || level > 8)
      return fail('The checklist indentation exceeds the supported native list levels.');
    return {
      kind: 'bullet',
      style: 'disc',
      level,
      dormantChecked: false,
      instance: { kind: 'new', key },
    };
  });
}
export function setManualChecklist(
  editor: DocxEditorInstance,
  checked: boolean,
  level: number,
  context: BlockCommandContext
) {
  return change(editor, context, () => ({ kind: 'checklist', checked, level }));
}
export function setManualListChecked(
  editor: DocxEditorInstance,
  checked: boolean,
  context: BlockCommandContext
) {
  return change(editor, context, (state) =>
    state.checkbox
      ? { kind: 'checklist', level: state.checkbox.level, checked }
      : fail('Select a checklist item.')
  );
}
export function adjustManualListLevel(
  editor: DocxEditorInstance,
  direction: 'increase' | 'decrease',
  context: BlockCommandContext
) {
  if (!editor.surface) return fail('The native editor is not attached.');
  editor.surface.flushPendingInput();
  const selected = selection(editor).map((id) => ({
    id,
    state: manualListState(editor, context, id),
  }));
  if (selected.some((item) => !item.state)) return fail('The selected list source is unavailable.');
  const surface = editor.surface,
    revision = readEditorRevision(editor, context.documentId, context.epoch);
  if (!revision) return fail('The native document is not ready.');
  const pkg = surface.session.currentPackage(),
    part = surface.session.partFor(surface.storyScope()) ?? surface.session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('List controls require the body story.');
  const operations: CanonicalEdit[] = [];
  for (const { id, state } of selected) {
    const current = state!;
    if (!current.checkbox && !current.numbering) {
      const patch = paragraphIndentPatch(editor, id, direction);
      if ('success' in patch) return patch;
      const planned = planParagraphFormat(
        { part, paragraphId: id, revision },
        { properties: [patch] }
      );
      if (!planned.success) return planned;
      operations.push(...planned.plan.operations);
      continue;
    }
    const old = current.checkbox?.level ?? current.numbering!.level,
      next = old + (direction === 'increase' ? 1 : -1),
      level = Math.min(8, next);
    let target: ListTarget;
    if (next < 0) target = { kind: 'none' };
    else if (current.checkbox)
      target = { kind: 'checklist', checked: current.checkbox.checked, level };
    else {
      const marker = manualListMarker(editor, id);
      if (!marker) return fail('The native list marker is unavailable.');
      target = {
        kind: marker.numFmt === 'bullet' ? 'bullet' : 'ordered',
        level,
        instance: { kind: 'existing', numId: current.numbering!.numId },
      };
    }
    const planned = planListMembership({ pkg, partName: part.name, revision }, [
      { paragraphId: id, target },
    ]);
    if (!planned.success) return planned;
    operations.push(...planned.plan.operations);
  }
  const planned = createEditPlan(revision, operations);
  if (!planned.success) return planned;
  if (!operations.length) return { success: true as const, committed: false, revision };
  return applyManualListPlan(editor, planned.plan, context);
}
export function manualListMarker(editor: DocxEditorInstance, paragraphId: string) {
  return editor.surface
    ?.layout()
    .pages.flatMap((page) => paragraphFragmentsOf(page))
    .find((p) => p.paragraphId === paragraphId)?.marker;
}
export function createListToolbar(editor: DocxEditorInstance, context: BlockCommandContext) {
  const s = editor.surface;
  if (!s) throw new Error('The native editor is not attached.');
  const doc = s.hostElement().ownerDocument,
    element = doc.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Lists');
  const style = doc.createElement('select');
  style.setAttribute('aria-label', 'List style');
  for (const value of [...LIST_STYLES, 'todo', 'none']) {
    const option = doc.createElement('option');
    option.value = value;
    option.textContent = value;
    style.append(option);
  }
  element.append(style);
  const output = doc.createElement('output');
  output.setAttribute('role', 'status');
  const buttons = [
    [
      'Apply list',
      () => setManualList(editor, style.value as ListStyle | 'todo' | 'none', context),
    ],
    [
      'Toggle list',
      () =>
        style.value === 'none'
          ? setManualList(editor, 'none', context)
          : toggleManualList(editor, style.value as ListStyle | 'todo', context),
    ],
    ['Insert bullets', () => insertManualList(editor, 'disc', context)],
    ['Insert numbers', () => insertManualList(editor, 'decimal', context)],
    ['Insert checklist', () => insertManualList(editor, 'todo', context)],
    ['Check item', () => setManualListChecked(editor, true, context)],
    ['Uncheck item', () => setManualListChecked(editor, false, context)],
    ['Indent checklist', () => indentManualChecklist(editor, context)],
    ['Indent list', () => adjustManualListLevel(editor, 'increase', context)],
    ['Outdent list', () => adjustManualListLevel(editor, 'decrease', context)],
  ] as const;
  for (const [label, run] of buttons) {
    const b = doc.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    b.addEventListener('click', () => {
      const result = run();
      output.textContent = result.success ? '' : result.error;
      if (result.success) editor.surface?.focus();
    });
    element.append(b);
  }
  element.append(output);
  const update = () => {
    const enabled = editor.snapshot().editable;
    style.disabled = !enabled;
    for (const b of Array.from(element.querySelectorAll('button'))) b.disabled = !enabled;
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
