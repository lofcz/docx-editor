import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { paragraphFragmentsOf, paragraphsInCells } from '@docx-editor.dev/core/layout';
import { directParagraphProperties } from '@docx-editor.dev/core/store';
import { applyDirectPlan } from '../ai/apply-direct';
import { planParagraphFormat, type ParagraphPropertyPatch } from '../diff/paragraph-format';
import { createEditPlan, type CanonicalEdit } from '../diff/plan';
import type { BlockCommandContext } from './block-type';
import type { NativeEditFailure } from '../contracts/document';

export const PARAGRAPH_ALIGNMENTS = ['start', 'left', 'center', 'right', 'end', 'justify'] as const;
export type ParagraphAlignment = (typeof PARAGRAPH_ALIGNMENTS)[number];
/** Production IndentKit uses 24 CSS pixels, or 18 points. */
export const PARAGRAPH_INDENT_STEP_TWIPS = 360;
const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});

function positionChange(
  editor: DocxEditorInstance,
  context: BlockCommandContext,
  change: (paragraphId: string) => ParagraphPropertyPatch | NativeEditFailure
) {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  const pkg = s.session.currentPackage(),
    part = s.session.part();
  if (part.name !== pkg.mainDocumentPart) return fail('Paragraph controls require the body story.');
  const { anchor, head } = s.state().selection,
    order = s.session.paragraphIds();
  const a = order.indexOf(anchor.paragraphId),
    h = order.indexOf(head.paragraphId);
  if (a < 0 || h < 0) return fail('The native selection has no paragraph.');
  const end = a <= h ? head : anchor;
  const cells = editor.getTableCellSelection()?.cellIds;
  const selected = cells?.length
    ? paragraphsInCells(s.layout(), cells)
    : order.slice(Math.min(a, h), Math.max(a, h) + (a !== h && end.offset === 0 ? 0 : 1));
  const revision = { ...context, revision: s.session.packageRevision() },
    operations: CanonicalEdit[] = [];
  for (const paragraphId of selected) {
    const patch = change(paragraphId);
    if ('success' in patch) return patch;
    const planned = planParagraphFormat({ part, paragraphId, revision }, { properties: [patch] });
    if (!planned.success) return planned;
    operations.push(...planned.plan.operations);
  }
  const plan = createEditPlan(revision, operations);
  return plan.success ? applyDirectPlan(editor, plan.plan, context) : plan;
}

/** Physical edges resolve per paragraph direction; start/end retain their logical Word values. */
export function setParagraphAlignment(
  editor: DocxEditorInstance,
  alignment: ParagraphAlignment,
  context: BlockCommandContext
) {
  if (!PARAGRAPH_ALIGNMENTS.includes(alignment)) return fail('Choose an application alignment.');
  return positionChange(editor, context, (id) => {
    const fragment = editor
      .surface!.layout()
      .pages.flatMap((p) => paragraphFragmentsOf(p))
      .find((p) => p.range.paragraphId === id);
    if (!fragment) return fail('The selected paragraph has no native layout.');
    const bidi = fragment.props.find((p) => p.localName === 'bidi');
    const rtl = !!bidi && !['0', 'false', 'off'].includes(bidi.attributes?.val ?? '');
    const value =
      alignment === 'justify'
        ? 'both'
        : rtl && alignment === 'left'
          ? 'right'
          : rtl && alignment === 'right'
            ? 'left'
            : alignment;
    return { localName: 'jc', attributes: { val: value } };
  });
}

/** Shift each paragraph independently without changing its numbering or first-line/hanging settings. */
export function changeParagraphIndent(
  editor: DocxEditorInstance,
  direction: 'increase' | 'decrease',
  context: BlockCommandContext
) {
  if (direction !== 'increase' && direction !== 'decrease')
    return fail('Choose an indentation direction.');
  return positionChange(editor, context, (id) => {
    return paragraphIndentPatch(editor, id, direction);
  });
}

/** Resolve each paragraph independently from its effective native indentation. */
export function paragraphIndentPatch(
  editor: DocxEditorInstance,
  id: string,
  direction: 'increase' | 'decrease'
): ParagraphPropertyPatch | NativeEditFailure {
  const s = editor.surface!;
  const fragment = s
    .layout()
    .pages.flatMap((p) => paragraphFragmentsOf(p))
    .find((p) => p.range.paragraphId === id);
  if (!fragment) return fail('The selected paragraph has no native layout.');
  const current = Math.round(fragment.indent.left * 20);
  const left = Math.max(
    0,
    current + (direction === 'increase' ? 1 : -1) * PARAGRAPH_INDENT_STEP_TWIPS
  );
  const authored = directParagraphProperties(s.session.part(), id).find(
    (p) => p.localName === 'ind'
  );
  // A character-based leading indent would override the new twip value in Word.
  return {
    localName: 'ind',
    attributes: {
      left: String(left),
      leftChars: null,
      ...(authored?.attributes?.start !== undefined ? { start: null, startChars: null } : {}),
    },
  };
}

export function createParagraphPositionToolbar(
  editor: DocxEditorInstance,
  context: BlockCommandContext
) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Paragraph position');
  const actions = [...PARAGRAPH_ALIGNMENTS, 'increase', 'decrease'] as const;
  const buttons = actions.map((action) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent =
      action === 'increase'
        ? 'Increase indent'
        : action === 'decrease'
          ? 'Decrease indent'
          : `Align ${action}`;
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      const result =
        action === 'increase' || action === 'decrease'
          ? changeParagraphIndent(editor, action, context)
          : setParagraphAlignment(editor, action, context);
      if (result.success) surface.focus();
    });
    element.append(button);
    return button;
  });
  const update = () => {
    const allowed = editor.can({ type: 'setAlignment', align: 'left' });
    for (const button of buttons) {
      button.disabled = !allowed.ok;
      button.title = allowed.ok ? '' : allowed.reason;
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
