import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { BASIC_MARKS, basicMarkState, toggleBasicMark, type BasicMark } from './marks';

const LABELS: Readonly<Record<BasicMark, string>> = {
  bold: 'Bold',
  italic: 'Italic',
  underline: 'Underline',
  strikethrough: 'Strikethrough',
  subscript: 'Subscript',
  superscript: 'Superscript',
  code: 'Inline code',
};

/** Dispose the subscription when the host removes this toolbar. */
export function createMarksToolbar(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Text formatting');
  const buttons = BASIC_MARKS.map((mark) => {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = LABELS[mark];
    button.addEventListener('mousedown', (event) => {
      if (event.button === 0) event.preventDefault();
    });
    button.addEventListener('click', () => {
      if (toggleBasicMark(editor, mark).ok) surface.focus();
    });
    element.append(button);
    return { button, mark };
  });
  const update = () => {
    for (const { button, mark } of buttons) {
      const commandMark = mark === 'code' ? 'bold' : mark === 'strikethrough' ? 'strike' : mark;
      const enabled = editor.can({ type: 'toggleMark', mark: commandMark });
      button.disabled = !enabled.ok;
      button.title = enabled.ok ? '' : enabled.reason;
      const state = basicMarkState(editor, mark);
      button.setAttribute('aria-pressed', state === 'mixed' ? 'mixed' : String(state === 'on'));
    }
  };
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  const dispose = () => stops.forEach((stop) => stop());
  update();
  return { element, dispose };
}

/** Application shortcuts take precedence over Word's paragraph shortcuts on the native surface. */
export function installMarkShortcuts(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const host = surface.hostElement();
  const keydown = (event: KeyboardEvent) => {
    const target = event.target as Element | null;
    if (
      !target?.closest?.('[contenteditable="true"]') ||
      target.closest('[data-docx-block-ui]') ||
      event.isComposing ||
      event.altKey ||
      event.getModifierState('AltGraph') ||
      !(event.ctrlKey || event.metaKey)
    )
      return;
    const key = event.key.toLowerCase();
    const shortcuts: Readonly<Record<string, BasicMark>> = {
      b: 'bold',
      i: 'italic',
      u: 'underline',
      e: 'code',
      ',': 'subscript',
      '.': 'superscript',
    };
    const mark = event.shiftKey ? (key === 'x' ? 'strikethrough' : undefined) : shortcuts[key];
    if (!mark) return;
    // Consume recognized chords even when refused, so browser commands cannot write around the gate.
    event.preventDefault();
    event.stopImmediatePropagation();
    toggleBasicMark(editor, mark);
  };
  host.addEventListener('keydown', keydown, true);
  return () => host.removeEventListener('keydown', keydown, true);
}
