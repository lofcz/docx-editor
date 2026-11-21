// When presses on the pages count as one multi-click: the pointer's word and paragraph
// selection, and the double press that enters a header or footer text box, read the same rule.

/** How long after a click a second one still counts as a double. */
export const MULTI_CLICK_MS = 500;
/** How far it may move and still count — a click, not a tiny drag. */
export const MULTI_CLICK_SLOP_PX = 4;

/** Count physical presses even when painting delays their handlers. */
export function createPressCounter(): (event: PointerEvent) => number {
  let previous: { time: number; x: number; y: number; view: Window | null; count: number } | null =
    null;
  return (event) => {
    const { timeStamp: time, clientX: x, clientY: y, view, detail } = event;
    // Pointer events in Chromium carry zero detail. Their timestamps still describe when
    // the device pressed. Comparing delivery times can turn a queued double click into two carets.
    const inferred =
      previous !== null &&
      previous.view === view &&
      time >= previous.time &&
      time - previous.time <= MULTI_CLICK_MS &&
      Math.abs(x - previous.x) <= MULTI_CLICK_SLOP_PX &&
      Math.abs(y - previous.y) <= MULTI_CLICK_SLOP_PX
        ? previous.count + 1
        : 1;
    const count = detail > 0 ? detail : inferred;
    previous = { time, x, y, view, count };
    return count;
  };
}
