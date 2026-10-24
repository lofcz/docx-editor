import { beginHistoryGroupFrames, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type { ExecResult } from '@docx-editor.dev/core/contracts/editor';
export const FONT_SIZE_PIXELS = [8, 9, 10, 12, 14, 16, 18, 24, 30, 36, 48, 60, 72, 96] as const;
export const PACKAGED_FONT_FAMILIES = [
  'Calibri',
  'Cambria',
  'Arial',
  'Times New Roman',
  'Courier New',
  'Century Gothic',
] as const;
const failed = (reason: string): ExecResult => ({ ok: false, code: 'invalidArgs', reason });
export function manualFontState(editor: DocxEditorInstance) {
  const f = editor.surface?.formatting();
  return {
    family: f?.fontFamily ?? null,
    sizePixels:
      f?.fontSizeHalfPoints === null || f?.fontSizeHalfPoints === undefined
        ? null
        : f.fontSizeHalfPoints / 1.5,
  };
}
/** Native family edits activate the editor's configured packagedFonts/application resolver. */
export function setManualFontFamily(editor: DocxEditorInstance, family: string | null): ExecResult {
  const s = editor.surface;
  if (!s) return failed('The native editor is not attached.');
  s.flushPendingInput();
  const allowed = editor.can({
    type: 'setMarkAttr',
    mark: 'fontFamily',
    attr: 'family',
    value: family ?? 'Calibri',
  });
  if (!allowed.ok) return { ok: false, code: 'locked', reason: allowed.reason };
  const revision = s.session.packageRevision(),
    reason = s.setRunProperty(
      'rFonts',
      family === null ? undefined : { ascii: family, hAnsi: family }
    );
  return reason ? failed(reason) : { ok: true, changed: revision !== s.session.packageRevision() };
}
/** Preserve the captured CSS-pixel value exactly; odd pixels require Word universal measures. */
export function setManualFontSize(editor: DocxEditorInstance, pixels: number | null): ExecResult {
  if (pixels !== null && (!Number.isFinite(pixels) || pixels < 4 / 3 || pixels > 2184))
    return failed('Choose a font size from 1.334 to 2184 pixels.');
  const s = editor.surface;
  if (!s) return failed('The native editor is not attached.');
  s.flushPendingInput();
  const allowed = editor.can({ type: 'setMarkAttr', mark: 'fontSize', attr: 'val', value: 22 });
  if (!allowed.ok) return { ok: false, code: 'locked', reason: allowed.reason };
  const value =
    pixels === null
      ? undefined
      : Number.isInteger(pixels * 1.5)
        ? String(pixels * 1.5)
        : `${pixels * 0.75}pt`;
  const group = editor.beginHistoryGroup(),
    frames = beginHistoryGroupFrames(editor, group);
  if (!frames) {
    group.end();
    return failed('The font change cannot start.');
  }
  const revision = s.session.packageRevision();
  try {
    for (const property of ['sz', 'szCs']) {
      const reason = frames.run(() =>
        s.setRunProperty(property, value === undefined ? undefined : { val: value })
      );
      if (reason) return failed(reason);
    }
    return { ok: true, changed: revision !== s.session.packageRevision() };
  } finally {
    frames.finish();
    group.end();
  }
}
export function createFontToolbar(
  editor: DocxEditorInstance,
  families: readonly string[] = PACKAGED_FONT_FAMILIES
) {
  const surface = editor.surface;
  if (!surface) throw new Error('The native editor is not attached.');
  const element = surface.hostElement().ownerDocument.createElement('div');
  element.setAttribute('role', 'toolbar');
  element.setAttribute('aria-label', 'Fonts');
  const family = element.ownerDocument.createElement('select');
  family.setAttribute('aria-label', 'Font family');
  for (const name of ['', ...families]) {
    const option = element.ownerDocument.createElement('option');
    option.value = name;
    option.textContent = name || 'Mixed / inherited';
    family.append(option);
  }
  const size = element.ownerDocument.createElement('input');
  size.type = 'number';
  size.step = '1';
  size.setAttribute('aria-label', 'Font size (px)');
  const sizes = element.ownerDocument.createElement('datalist');
  sizes.id = `font-sizes-${editor.mountGeneration}-${Math.random().toString(36).slice(2)}`;
  for (const value of FONT_SIZE_PIXELS) {
    const option = element.ownerDocument.createElement('option');
    option.value = String(value);
    sizes.append(option);
  }
  size.setAttribute('list', sizes.id);
  element.append(sizes);
  const status = element.ownerDocument.createElement('output');
  status.setAttribute('role', 'status');
  const run = (change: () => ExecResult) => {
    const result = change();
    status.textContent = result.ok ? '' : result.reason;
    if (result.ok) editor.surface?.focus();
  };
  family.addEventListener('change', () =>
    run(() => setManualFontFamily(editor, family.value || null))
  );
  const buttons = [
    ['Apply font size', () => setManualFontSize(editor, Number(size.value))],
    ['Clear font size', () => setManualFontSize(editor, null)],
    ['Clear font family', () => setManualFontFamily(editor, null)],
  ] as const;
  element.append(family, size);
  for (const [label, change] of buttons) {
    const button = element.ownerDocument.createElement('button');
    button.type = 'button';
    button.textContent = label;
    button.addEventListener('mousedown', (e) => {
      if (e.button === 0) e.preventDefault();
    });
    button.addEventListener('click', () => run(change));
    element.append(button);
  }
  const update = () => {
    const f = manualFontState(editor);
    if (f.family && !Array.from(family.options).some((option) => option.value === f.family)) {
      const option = element.ownerDocument.createElement('option');
      option.value = f.family;
      option.textContent = f.family;
      family.append(option);
    }
    family.value = f.family ?? '';
    size.value = f.sizePixels === null ? '' : String(f.sizePixels);
    const allowed = editor.can({ type: 'setMarkAttr', mark: 'fontSize', attr: 'val', value: 22 });
    family.disabled = !allowed.ok;
    size.disabled = !allowed.ok;
    for (const button of Array.from(element.querySelectorAll('button')))
      button.disabled = !allowed.ok;
  };
  element.append(status);
  const stops = [editor.on('change', update), editor.on('selectionChange', update)];
  update();
  return { element, dispose: () => stops.forEach((stop) => stop()) };
}
