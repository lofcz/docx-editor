// ECMA-376 Part 1 §17.3.3.31 (w:t, CT_Text). Both text nodes paint the same
// literal run. The browser can keep its caret at the start of the unedited suffix.
const suffixes = new WeakMap<Element, Text>();

export function typingSuffix(run: Element): Text | undefined {
  const suffix = suffixes.get(run);
  if (
    suffix &&
    suffix.parentNode === run &&
    run.childNodes.length === 2 &&
    run.lastChild === suffix &&
    run.firstChild?.nodeType === 3
  )
    return suffix;
  return undefined;
}

function patchText(node: Text, value: string, split?: HTMLElement): void {
  const before = node.data;
  if (before === value) return;
  let start = 0;
  let end = 0;
  const common = Math.min(before.length, value.length);
  while (start < common && before[start] === value[start]) start++;
  while (end < common - start && before[before.length - end - 1] === value[value.length - end - 1])
    end++;
  const removed = before.length - start - end;
  const inserted = value.slice(start, value.length - end);
  if (split && !removed && end && inserted && !split.hasAttribute('data-docx-field')) {
    const selection = node.ownerDocument.getSelection();
    if (selection?.isCollapsed && selection.focusNode === node && selection.focusOffset === start) {
      const suffix = node.splitText(start);
      node.appendData(inserted);
      suffixes.set(split, suffix);
      return;
    }
  }
  node.replaceData(start, removed, inserted);
}

export function updateTypingText(run: HTMLElement, value: string): void {
  const before = run.firstChild as Text;
  const suffix = typingSuffix(run);
  if (!suffix) return patchText(before, value, run);
  if (value.endsWith(suffix.data)) {
    patchText(before, value.slice(0, value.length - suffix.length));
    return;
  }
  suffixes.delete(run);
  suffix.remove();
  patchText(before, value);
}
