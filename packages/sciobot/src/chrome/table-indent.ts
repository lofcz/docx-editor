import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { WML_NAMESPACE_URI, type OoxmlNode } from '@docx-editor.dev/core/store';
import {
  applyManualTableOps,
  canManualTableOps,
  manualTableContext,
  tableRefusal,
} from './table-native';

const attr = (node: OoxmlNode, name: string) =>
  node.kind === 'textValue'
    ? undefined
    : node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
        ?.value;
/** Read only the authored indent. An omitted value can inherit from a Word table style. */
export function manualTableIndent(editor: DocxEditorInstance): number | null {
  const context = manualTableContext(editor);
  if (!context) return null;
  const props = context.table.children.find((n) => n.kind === 'tableProperties');
  if (!props) return null;
  const indent = props.children.find(
    (n) => n.namespaceUri === WML_NAMESPACE_URI && n.localName === 'tblInd'
  );
  if (!indent || attr(indent, 'type') !== 'dxa') return null;
  const twips = Number(attr(indent, 'w'));
  return Number.isFinite(twips) ? twips / 15 : null;
}
function operation(editor: DocxEditorInstance, pixels: number) {
  const context = manualTableContext(editor),
    twips = pixels * 15;
  if (!context || !Number.isInteger(twips) || Math.abs(twips) > 31680) return null;
  return { op: 'setTableProperties' as const, tableId: context.table.id, indentTwips: twips };
}
export function setManualTableIndent(editor: DocxEditorInstance, pixels: number) {
  editor.surface?.flushPendingInput();
  const op = operation(editor, pixels);
  return op
    ? applyManualTableOps(editor, [op])
    : tableRefusal('Select a table and use an exact indent from -2112 to 2112 pixels.');
}
export function createTableIndentToolbar(editor: DocxEditorInstance) {
  const s = editor.surface;
  if (!s) throw new Error('The native editor is not attached.');
  const doc = s.hostElement().ownerDocument,
    element = doc.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Table indentation');
  const input = doc.createElement('input');
  input.type = 'number';
  input.step = '1';
  input.setAttribute('aria-label', 'Table left margin (px)');
  const status = doc.createElement('output');
  status.setAttribute('role', 'status');
  element.append(input);
  for (const [label, value] of [
    ['Apply table margin', () => Number(input.value)],
    ['Reset table margin', () => 0],
  ] as const) {
    const button = doc.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    button.addEventListener('click', () => {
      const result = setManualTableIndent(editor, value());
      status.textContent = result.ok ? '' : result.reason;
      if (result.ok) editor.surface?.focus();
    });
    element.append(button);
  }
  element.append(status);
  const update = () => {
    const pixels = manualTableIndent(editor);
    input.value = pixels === null ? '' : String(pixels);
    const op = operation(editor, pixels ?? 0),
      enabled = !!op && canManualTableOps(editor, [op]).ok;
    input.disabled = !enabled;
    for (const button of Array.from(element.querySelectorAll('button'))) button.disabled = !enabled;
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
