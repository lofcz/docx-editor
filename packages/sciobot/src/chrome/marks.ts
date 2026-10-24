import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
import { spansInSelection, spansInCells, paragraphsInCells } from '@docx-editor.dev/core/layout';
import { paragraphTextOf, type OoxmlProperty, type TreeDocOp } from '@docx-editor.dev/core/store';
import {
  inlineCodeCharacterStyles,
  codeStyleBase,
  inlineCodeOps,
  inlineCodeStyleId,
  inlineCodeStyleState,
  withInlineCodeStyle,
} from '../native/inline-code';

export const BASIC_MARKS = [
  'bold',
  'italic',
  'underline',
  'strikethrough',
  'subscript',
  'superscript',
  'code',
] as const;
export type BasicMark = (typeof BASIC_MARKS)[number];
export type MarkState = 'on' | 'off' | 'mixed';

const nativeMark = (mark: Exclude<BasicMark, 'code'>) =>
  mark === 'strikethrough' ? 'strike' : mark;
const refused = (reason: string): ExecResult => ({ ok: false, code: 'locked', reason });

/** Read the native cascade, including locally armed caret properties. */
export function basicMarkState(editor: DocxEditorInstance, mark: BasicMark): MarkState {
  const surface = editor.surface;
  if (!surface) return 'off';
  const { selection } = surface.state();
  const pendingFormat = surface.state().pendingFormat as readonly OoxmlProperty[] | null;
  const cells = editor.getTableCellSelection()?.cellIds;
  const collapsed =
    !cells?.length &&
    selection.anchor.paragraphId === selection.head.paragraphId &&
    selection.anchor.offset === selection.head.offset;
  if (collapsed && mark !== 'code') return surface.formatting()[mark] ? 'on' : 'off';
  const codeStyles = new Set(inlineCodeCharacterStyles(surface.session.currentPackage()));
  const pendingStyle = collapsed && pendingFormat?.find((p) => p.localName === 'rStyle');
  if (mark === 'code' && pendingStyle)
    return codeStyles.has(pendingStyle.attributes?.val ?? '') ? 'on' : 'off';
  const spans = cells?.length
    ? spansInCells(surface.layout(), cells)
    : spansInSelection(surface.layout(), selection, surface.session.paragraphIds());
  const values = spans.map((span) => {
    switch (mark) {
      case 'code':
        return codeStyles.has(
          span.props.find((p) => p.localName === 'rStyle')?.attributes?.val ?? ''
        );
      case 'underline':
        return span.style.underline !== null;
      case 'strikethrough':
        return span.style.strike;
      case 'subscript':
      case 'superscript':
        return span.style.verticalAlign === mark;
      default:
        return span.style[mark];
    }
  });
  return values.length && values.every(Boolean) ? 'on' : values.some(Boolean) ? 'mixed' : 'off';
}

/** Use native commands for six marks and explicit Word character styles for code. */
export function toggleBasicMark(editor: DocxEditorInstance, mark: BasicMark): ExecResult {
  const surface = editor.surface;
  if (!surface) return refused('The native editor is not attached.');
  surface.flushPendingInput();
  if (mark !== 'code') return editor.exec({ type: 'toggleMark', mark: nativeMark(mark) });
  const allowed = editor.can({ type: 'toggleMark', mark: 'bold' });
  if (!allowed.ok) return refused(allowed.reason);
  const pkg = surface.session.currentPackage();
  const part = surface.session.part();
  if (part.name !== pkg.mainDocumentPart) return refused('Code controls require the body story.');
  const code = basicMarkState(editor, mark) !== 'on';
  const { selection } = surface.state();
  const pendingFormat = surface.state().pendingFormat as readonly OoxmlProperty[] | null;
  const order = surface.session.paragraphIds();
  const a = order.indexOf(selection.anchor.paragraphId),
    h = order.indexOf(selection.head.paragraphId);
  if (a < 0 || h < 0) return refused('The native selection has no paragraph.');
  const forward = a < h || (a === h && selection.anchor.offset <= selection.head.offset);
  const from = forward ? selection.anchor : selection.head;
  const to = forward ? selection.head : selection.anchor;
  const cells = editor.getTableCellSelection()?.cellIds;
  if (!cells?.length && a === h && from.offset === to.offset) {
    const spans = spansInSelection(surface.layout(), selection, order);
    const before =
      pendingFormat?.find((p) => p.localName === 'rStyle')?.attributes?.val ??
      spans[0]?.props.find((p) => p.localName === 'rStyle')?.attributes?.val;
    const codeStyles = inlineCodeCharacterStyles(pkg);
    const style = code
      ? inlineCodeStyleId(before && !codeStyles.includes(before) ? before : undefined)
      : before
        ? codeStyleBase(before)
        : undefined;
    if (code && typeof style === 'string' && inlineCodeStyleState(pkg, style) !== 'present')
      return refused('Prepare the document code styles before editing.');
    const error = surface.setRunProperty('rStyle', { val: style ?? '' });
    return error ? refused(error) : { ok: true, changed: false };
  }
  const ops: TreeDocOp[] = [];
  const styles = new Set(inlineCodeCharacterStyles(pkg));
  const view = surface.revisionDisplayMode();
  const displayMode =
    view === 'original' ? 'original' : view === 'all-markup' ? 'all-markup' : 'proposed';
  const selected = cells?.length
    ? paragraphsInCells(surface.layout(), cells)
    : order.slice(Math.min(a, h), Math.max(a, h) + 1);
  for (const id of selected) {
    const start = !cells?.length && id === from.paragraphId ? from.offset : 0;
    const end =
      !cells?.length && id === to.paragraphId
        ? to.offset
        : (paragraphTextOf(part, id)?.length ?? 0);
    if (start === end) continue;
    const planned = inlineCodeOps(pkg, { paragraphId: id, start, end, code }, displayMode);
    if (!planned.success) return refused(planned.error);
    ops.push(...planned.ops);
    for (const op of planned.ops)
      if (code && op.op === 'setRunProperties' && typeof op.characterStyle === 'string')
        styles.add(op.characterStyle);
  }
  if (!ops.length) return { ok: true, changed: false };
  const ids = [...styles];
  if (!withInlineCodeStyle(pkg, ids)) return refused('The document cannot gain the code styles.');
  const result = surface.applyAutomationOps(
    () => ops,
    { kind: 'body' },
    [
      (working) => {
        const prepared = withInlineCodeStyle(working, ids);
        if (!prepared) throw new Error('The document cannot gain the code styles.');
        return prepared;
      },
    ],
    false,
    true
  );
  return result.rejected ? refused(String(result.reason)) : { ok: true, changed: result.committed };
}
