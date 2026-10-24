import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';

/** Manual authors come from the host user. AI sessions supply attribution to their own executor. */
export function manualSuggestionState(editor: DocxEditorInstance) {
  const active = editor.surface?.editingMode() === 'suggest';
  const capability = editor.can({
    type: 'setEditingMode',
    mode: active ? 'editing' : 'suggesting',
  });
  return { active, author: editor.getConfiguredAuthor(), capability };
}
export function setManualSuggestionMode(editor: DocxEditorInstance, enabled: boolean): ExecResult {
  const surface = editor.surface;
  if (!surface)
    return { ok: false, code: 'invalidArgs', reason: 'The native editor is not attached.' };
  const command = {
    type: 'setEditingMode' as const,
    mode: enabled ? ('suggesting' as const) : ('editing' as const),
  };
  const allowed = editor.can(command);
  // A refused toggle must not leave a pending author request that later changes the mode.
  if (!allowed.ok) return { ok: false, code: 'locked', reason: allowed.reason };
  surface.flushPendingInput();
  if (editor.surface !== surface)
    return {
      ok: false,
      code: 'invalidArgs',
      reason: 'The document changed. Try the mode change again.',
    };
  return editor.exec(command);
}
export function createManualSuggestionToggle(editor: DocxEditorInstance) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const doc = surface.hostElement().ownerDocument,
    element = doc.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Manual suggestions');
  const button = doc.createElement('button');
  button.type = 'button';
  button.textContent = 'Suggestion edits';
  const status = doc.createElement('output');
  status.setAttribute('role', 'status');
  button.addEventListener('mousedown', (event) => {
    if (event.button === 0) event.preventDefault();
  });
  button.addEventListener('click', () => {
    const result = setManualSuggestionMode(editor, !manualSuggestionState(editor).active);
    status.textContent = result.ok ? '' : result.reason;
    if (result.ok) editor.surface?.focus();
  });
  const update = () => {
    const state = manualSuggestionState(editor);
    button.setAttribute('aria-pressed', String(state.active));
    button.disabled = !state.capability.ok;
    button.title = state.capability.ok
      ? state.active
        ? 'Turn off suggesting'
        : 'Suggestion edits'
      : state.capability.reason;
  };
  element.append(button, status);
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
