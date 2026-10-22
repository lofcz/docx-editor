/**
 * Widget form events stay with React controls. Native slot content keeps the editor's routing.
 *
 * React handles these events at its root container, before this listener runs. Stopping them
 * there keeps the editor's bubble-phase keymap, clipboard, composition, and pointer handlers
 * away from widget controls. The core `data-docx-host-ui` marker covers its capture-phase
 * readers, focus, and selection.
 */
export const WIDGET_INPUT_EVENTS = Object.freeze([
  'pointerdown',
  'pointerup',
  'mousedown',
  'mouseup',
  'click',
  'dblclick',
  'contextmenu',
  'selectstart',
  'beforeinput',
  'input',
  'keydown',
  'keyup',
  'keypress',
  'compositionstart',
  'compositionupdate',
  'compositionend',
  'paste',
  'copy',
  'cut',
] as const);

const WIDGET_UI = '[data-sciobot-block-ui]';

/** True when an event started inside widget furniture. Native slot content returns false. */
export function isWidgetInputEvent(event: Event): boolean {
  const target = event.target as Node | null;
  const element = target?.nodeType === 1 ? (target as Element) : target?.parentElement;
  return !!element?.closest(WIDGET_UI);
}

function stopWidgetEvent(event: Event): void {
  if (isWidgetInputEvent(event)) event.stopPropagation();
}

/** Attach after `createRoot(element)`, so React's own root listeners run first. */
export function isolateWidgetInput(element: HTMLElement): () => void {
  for (const type of WIDGET_INPUT_EVENTS) element.addEventListener(type, stopWidgetEvent);
  return () => {
    for (const type of WIDGET_INPUT_EVENTS) element.removeEventListener(type, stopWidgetEvent);
  };
}
