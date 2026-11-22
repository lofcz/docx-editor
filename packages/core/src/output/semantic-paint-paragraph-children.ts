// Only the painter owns the nodes recorded here. Hosts can attach controls to a paragraph.
const ownedChildren = new WeakMap<HTMLElement, readonly HTMLElement[]>();

export function recordParagraphChildren(root: HTMLElement): HTMLElement {
  const children = Array.from(root.children) as HTMLElement[];
  for (const child of children) {
    if (child.dataset.lineId) ownedChildren.set(child, Array.from(child.children) as HTMLElement[]);
  }
  ownedChildren.set(root, children);
  return root;
}

function copyAttributes(target: HTMLElement, source: HTMLElement): void {
  for (const name of target.getAttributeNames()) {
    if (!source.hasAttribute(name)) target.removeAttribute(name);
  }
  for (const attr of Array.from(source.attributes)) {
    if (target.getAttribute(attr.name) !== attr.value) target.setAttribute(attr.name, attr.value);
  }
}

function setRunText(node: Text, value: string): void {
  const before = node.data;
  if (before === value) return;
  let start = 0;
  let end = 0;
  const common = Math.min(before.length, value.length);
  while (start < common && before[start] === value[start]) start++;
  while (end < common - start && before[before.length - end - 1] === value[value.length - end - 1])
    end++;
  node.replaceData(start, before.length - start - end, value.slice(start, value.length - end));
}

function plainRun(node: HTMLElement | undefined): node is HTMLElement {
  return (
    node?.tagName === 'SPAN' &&
    node.classList.contains('layout-run-text') &&
    node.childNodes.length === 1 &&
    node.firstChild?.nodeType === 3
  );
}

function reconcileOwned(
  root: HTMLElement,
  fresh: HTMLElement,
  choose: (child: HTMLElement, index: number) => HTMLElement
): void {
  const old = ownedChildren.get(root)!;
  const next = Array.from(fresh.children, (child, index) => choose(child as HTMLElement, index));
  const kept = new Set(next);
  const replacing = new Set(old);
  // Remove retired furniture before ordering survivors. Moving a surviving line
  // across a retired sibling can disconnect a browser selection in that line.
  let tail: ChildNode | null = null;
  for (const child of old) if (child.parentNode === root) tail = child.nextSibling;
  for (const child of old) {
    if (!kept.has(child) && child.parentNode === root) child.remove();
  }
  let cursor = root.firstElementChild as HTMLElement | null;
  for (const child of next) {
    while (cursor && !replacing.has(cursor))
      cursor = cursor.nextElementSibling as HTMLElement | null;
    if (cursor === child) cursor = cursor.nextElementSibling as HTMLElement | null;
    else root.insertBefore(child, cursor ?? tail);
  }
  ownedChildren.set(root, next);
}

/** Keep text and its browser selection connected. Replace opaque paint at its own position. */
export function reconcileParagraphChildren(root: HTMLElement, fresh: HTMLElement): void {
  const old = ownedChildren.get(root)!;
  const lines = new Map(
    old.filter((child) => child.dataset.lineId).map((c) => [c.dataset.lineId, c])
  );
  reconcileOwned(root, fresh, (source) => {
    const line = lines.get(source.dataset.lineId);
    const runs = line && ownedChildren.get(line);
    if (!line || line.parentNode !== root || !runs) return source;
    copyAttributes(line, source);
    reconcileOwned(line, source, (child, i) => {
      const run = runs[i];
      if (!plainRun(run) || run.parentNode !== line || !plainRun(child)) return child;
      copyAttributes(run, child);
      setRunText(run.firstChild as Text, child.textContent ?? '');
      return run;
    });
    return line;
  });
}

const boxProperties = new Set(['top', 'left', 'width', 'height']);

export function canRetainPaintedParagraph(root: HTMLElement, fresh: HTMLElement): boolean {
  if (!ownedChildren.has(root) || root.attributes.length !== fresh.attributes.length) return false;
  for (const attr of Array.from(fresh.attributes)) {
    if (attr.name !== 'style' && root.getAttribute(attr.name) !== attr.value) return false;
  }
  // Flow can move or grow a paragraph across edits. Keep the element while
  // layout updates its box, including when a word wraps to a second line.
  if (root.style.cssText === fresh.style.cssText) return true;
  for (const style of [root.style, fresh.style]) {
    for (let i = 0; i < style.length; i++) {
      const property = style.item(i);
      if (
        !boxProperties.has(property) &&
        root.style.getPropertyValue(property) !== fresh.style.getPropertyValue(property)
      )
        return false;
    }
  }
  return true;
}

export function retainPaintedParagraph(
  root: HTMLElement | undefined,
  fresh: HTMLElement
): HTMLElement {
  if (!root || !canRetainPaintedParagraph(root, fresh)) return fresh;
  if (root.style.cssText !== fresh.style.cssText) root.style.cssText = fresh.style.cssText;
  reconcileParagraphChildren(root, fresh);
  return root;
}
