import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
import { spansInSelection, spansInCells } from '@docx-editor.dev/core/layout';
import type { OoxmlProperty } from '@docx-editor.dev/core/store';

export type ColorChannel = 'foreground' | 'background';
export interface ColorSelectionState {
  readonly value: string | null;
  readonly mixed: boolean;
}
const propertyName = (channel: ColorChannel) => (channel === 'foreground' ? 'color' : 'shd');
const highlights: Readonly<Record<string, string>> = {
  black: '000000',
  blue: '0000FF',
  cyan: '00FFFF',
  green: '00FF00',
  magenta: 'FF00FF',
  red: 'FF0000',
  yellow: 'FFFF00',
  white: 'FFFFFF',
  darkBlue: '000080',
  darkCyan: '008080',
  darkGreen: '008000',
  darkMagenta: '800080',
  darkRed: '800000',
  darkYellow: '808000',
  darkGray: '808080',
  lightGray: 'C0C0C0',
};

/** Hex literals are document data. Never accept CSS expressions or external resources. */
export function setTextColor(
  editor: DocxEditorInstance,
  channel: ColorChannel,
  color: string | null
): ExecResult {
  if (channel !== 'foreground' && channel !== 'background')
    return { ok: false, code: 'invalidArgs', reason: 'Choose a text color channel.' };
  const normalized =
    color === null || color === 'inherit' || color === 'transparent'
      ? null
      : color.replace(/^#/, '').toUpperCase();
  if (normalized !== null && !/^[0-9A-F]{6}$/.test(normalized))
    return { ok: false, code: 'invalidArgs', reason: 'Choose a six-digit RGB color or clear it.' };
  const surface = editor.surface;
  if (!surface) return { ok: false, code: 'locked', reason: 'The native editor is not attached.' };
  surface.flushPendingInput();
  const allowed = editor.can({ type: 'setMarkAttr', mark: 'color', attr: 'val', value: 'auto' });
  if (!allowed.ok) return { ok: false, code: 'locked', reason: allowed.reason };
  const revision = surface.session.packageRevision();
  const reason = surface.setRunProperty(
    propertyName(channel),
    normalized === null
      ? undefined
      : channel === 'foreground'
        ? { val: normalized }
        : { val: 'clear', color: 'auto', fill: normalized }
  );
  return reason
    ? { ok: false, code: 'locked', reason }
    : { ok: true, changed: revision !== surface.session.packageRevision() };
}

/** Resolved values drive the picker; reads never flatten the document's theme references. */
export function textColorState(
  editor: DocxEditorInstance,
  channel: ColorChannel
): ColorSelectionState {
  const surface = editor.surface;
  if (!surface) return { value: null, mixed: false };
  const state = surface.state(),
    cells = editor.getTableCellSelection()?.cellIds;
  const collapsed =
    !cells?.length &&
    state.selection.anchor.paragraphId === state.selection.head.paragraphId &&
    state.selection.anchor.offset === state.selection.head.offset;
  if (collapsed) {
    const pending = (state.pendingFormat as readonly OoxmlProperty[] | null)?.find(
      (p) => p.localName === propertyName(channel)
    );
    if (pending) {
      const hex = pending.attributes?.[channel === 'foreground' ? 'val' : 'fill'];
      return {
        value: hex && /^[0-9a-f]{6}$/i.test(hex) ? `#${hex.toUpperCase()}` : null,
        mixed: false,
      };
    }
  }
  const spans = cells?.length
    ? spansInCells(surface.layout(), cells)
    : spansInSelection(surface.layout(), state.selection, surface.session.paragraphIds());
  const values = spans.map((span) =>
    channel === 'foreground'
      ? span.style.color
      : (highlights[span.style.highlight ?? ''] ?? span.style.shading)
  );
  const normalize = (value: string | null | undefined) =>
    value ? `#${value.replace(/^#/, '').toUpperCase()}` : null;
  const value = normalize(values[0]),
    mixed = values.some((v) => normalize(v) !== value);
  return { value: mixed ? null : value, mixed };
}

/** A host may supply its production palette; selection remains native during pointer actions. */
export function createColorToolbar(editor: DocxEditorInstance, palette: readonly string[]) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Text colors');
  const controls = (['foreground', 'background'] as const).map((channel) => {
    const current = element.ownerDocument.createElement('output');
    current.setAttribute('aria-label', `Current ${channel}`);
    element.append(current);
    const buttons = [null, ...palette].map((color) => {
      const button = element.ownerDocument.createElement('button');
      button.type = 'button';
      button.textContent = `${channel} ${color ?? 'clear'}`;
      button.addEventListener('mousedown', (event) => {
        if (event.button === 0) event.preventDefault();
      });
      button.addEventListener('click', () => {
        if (setTextColor(editor, channel, color).ok) surface.focus();
      });
      element.append(button);
      return button;
    });
    return { current, buttons, channel };
  });
  const update = () => {
    const allowed = editor.can({ type: 'setMarkAttr', mark: 'color', attr: 'val', value: 'auto' });
    for (const { current, buttons, channel } of controls) {
      const state = textColorState(editor, channel);
      current.textContent = state.mixed ? 'Mixed' : (state.value ?? 'Default');
      for (const button of buttons) {
        button.disabled = !allowed.ok;
        button.title = allowed.ok ? '' : allowed.reason;
      }
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
