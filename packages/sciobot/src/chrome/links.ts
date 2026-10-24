import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult, SelectionPin } from '@docx-editor.dev/core/contracts/editor';
import { sanitizeHref } from '@docx-editor.dev/core/store';
const fail = (reason: string): ExecResult => ({ ok: false, code: 'invalidArgs', reason });
export function applyManualLink(
  editor: DocxEditorInstance,
  href: string,
  tooltip?: string
): ExecResult {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  const safe = sanitizeHref(href);
  if (!safe.ok || !safe.href) return fail('Choose a safe link target.');
  if (!editor.snapshot().editable)
    return { ok: false, code: 'locked', reason: 'The document is open for viewing.' };
  const existing = s.hyperlinks.linkAtCaret(),
    selection = s.state().selection;
  if (
    existing &&
    existing.authored === safe.href &&
    (tooltip === undefined || tooltip === existing.tooltip) &&
    [selection.anchor, selection.head].every(
      (p) =>
        p.paragraphId === existing.paragraphId &&
        p.offset >= existing.start &&
        p.offset <= existing.end
    )
  )
    return { ok: true, changed: false };
  const changed = s.hyperlinks.applyHyperlink({
    ...(safe.href.startsWith('#') ? { anchor: safe.href.slice(1) } : { url: safe.href }),
    ...(tooltip === undefined ? {} : { tooltip }),
  });
  return changed
    ? { ok: true, changed: true }
    : fail('The native editor refused this link target.');
}
export function removeManualLink(editor: DocxEditorInstance): ExecResult {
  const s = editor.surface;
  if (!s) return fail('The native editor is not attached.');
  s.flushPendingInput();
  return s.hyperlinks.removeHyperlink()
    ? { ok: true, changed: true }
    : fail('Select a link to remove.');
}
export function activateManualLink(editor: DocxEditorInstance, linkId?: string): boolean {
  const s = editor.surface;
  if (!s) return false;
  const link = linkId
    ? s.hyperlinks.linkById(linkId)
    : (s.hyperlinks.linkAtCaret() ?? s.hyperlinks.fieldLinkAtCaret());
  if (!link?.href) return false;
  return link.kind === 'internal'
    ? s.navigation.goToBookmark(link.anchor ?? link.href.slice(1))
    : s.navigation.openExternal(link.href);
}
/** Keep native selection pins live while focus moves into the link form. */
export function createLinkToolbar(editor: DocxEditorInstance) {
  const s = editor.surface;
  if (!s) throw new Error('The native editor is not attached.');
  const doc = s.hostElement().ownerDocument;
  const element = doc.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Links');
  const form = doc.createElement('div');
  form.setAttribute('role', 'dialog');
  form.setAttribute('aria-label', 'Edit link');
  form.hidden = true;
  const input = doc.createElement('input');
  input.type = 'text';
  input.setAttribute('aria-label', 'Link target');
  const status = doc.createElement('output');
  status.setAttribute('role', 'status');
  let pin: SelectionPin | null = null;
  let openedSurface = editor.surface;
  const close = () => {
    form.hidden = true;
    if (pin) editor.releaseSelection(pin);
    pin = null;
    editor.surface?.focus();
  };
  const open = (initialHref?: string) => {
    const s = editor.surface;
    if (!s) return;
    if (pin) editor.releaseSelection(pin);
    s.flushPendingInput();
    openedSurface = s;
    pin = editor.retainSelection();
    input.value = initialHref ?? s.hyperlinks.linkAtCaret()?.authored ?? '';
    form.hidden = false;
    input.focus();
  };
  const button = (label: string, action: () => void, parent: HTMLElement = element) => {
    const b = doc.createElement('button');
    b.type = 'button';
    b.textContent = label;
    b.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    b.addEventListener('click', () => action());
    parent.append(b);
    return b;
  };
  button('Edit link', open);
  form.append(input);
  button(
    'Apply link',
    () => {
      if (openedSurface !== editor.surface) {
        close();
        status.textContent = 'The document changed. Open the link dialog again.';
        return;
      }
      const result = applyManualLink(editor, input.value);
      status.textContent = result.ok ? '' : result.reason;
      if (result.ok) close();
    },
    form
  );
  button(
    'Remove link',
    () => {
      if (openedSurface !== editor.surface) {
        close();
        status.textContent = 'The document changed. Open the link dialog again.';
        return;
      }
      const result = removeManualLink(editor);
      status.textContent = result.ok ? '' : result.reason;
      if (result.ok) close();
    },
    form
  );
  button('Cancel', close, form);
  button('Open link', () => {
    if (!activateManualLink(editor)) status.textContent = 'This link cannot be opened.';
  });
  element.append(form, status);
  const stopChrome = editor.setHyperlinkChrome({
    onRequest: open,
    onPopover: ({ link }) => {
      const surface = editor.surface;
      if (!surface) return;
      surface.setSelection({
        anchor: { paragraphId: link.paragraphId, offset: link.start },
        head: { paragraphId: link.paragraphId, offset: link.end },
      });
      open(link.authored);
    },
  });
  const keydown = (event: KeyboardEvent) => {
    if (event.key === 'Escape' && !form.hidden) {
      event.preventDefault();
      close();
    }
  };
  form.addEventListener('keydown', keydown);
  return {
    element,
    open,
    close,
    dispose: () => {
      stopChrome();
      form.removeEventListener('keydown', keydown);
      if (pin) editor.releaseSelection(pin);
      pin = null;
    },
  };
}
