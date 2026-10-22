export type JsonValue = null | boolean | number | string | readonly JsonValue[] | JsonObject;
export interface JsonObject {
  readonly [key: string]: JsonValue;
}

export interface JsonDataLimits {
  readonly maxDepth: number;
  readonly maxNodes: number;
}

/** Validate before JSON.stringify can silently drop values or invoke toJSON. */
export function jsonDataProblem(value: unknown, limits: JsonDataLimits): string | null {
  const stack: { value: unknown; depth: number; exit?: true }[] = [{ value, depth: 0 }];
  const ancestors = new Set<object>();
  let visited = 0;
  while (stack.length) {
    const entry = stack.pop()!;
    if (entry.exit) {
      ancestors.delete(entry.value as object);
      continue;
    }
    if (++visited > limits.maxNodes) return 'node-limit';
    if (entry.depth > limits.maxDepth) return 'depth-limit';
    const node = entry.value;
    if (node === null || typeof node === 'string' || typeof node === 'boolean') continue;
    if (typeof node === 'number' && Number.isFinite(node)) continue;
    if (typeof node !== 'object') return 'non-json-value';
    const prototype = Object.getPrototypeOf(node);
    if (prototype !== null && prototype !== Object.prototype && prototype !== Array.prototype)
      return 'non-json-object';
    if (ancestors.has(node)) return 'cyclic-payload';
    if (Object.getOwnPropertySymbols(node).length) return 'non-json-property';
    if (Array.isArray(node) && node.length > limits.maxNodes) return 'node-limit';
    const descriptors = Object.getOwnPropertyDescriptors(node);
    const keys = Object.keys(descriptors).filter((key) => key !== 'length' || !Array.isArray(node));
    if (keys.length > limits.maxNodes) return 'node-limit';
    if (
      Array.isArray(node) &&
      (keys.length !== node.length || keys.some((key, index) => key !== String(index)))
    )
      return 'sparse-or-extended-array';
    ancestors.add(node);
    stack.push({ value: node, depth: entry.depth, exit: true });
    for (const key of keys) {
      const descriptor = descriptors[key]!;
      if (!('value' in descriptor) || !descriptor.enumerable) return 'non-json-property';
      stack.push({ value: descriptor.value, depth: entry.depth + 1 });
    }
  }
  return null;
}

/** Detached bounded data copy. Check string sizes before JSON escaping can allocate a large result. */
export function copyBoundedJson(
  value: unknown,
  limits: JsonDataLimits & { readonly maxBytes: number }
):
  | { readonly ok: true; readonly value: JsonValue }
  | { readonly ok: false; readonly reason: string } {
  const problem = jsonDataProblem(value, limits);
  if (problem) return { ok: false, reason: problem };
  const stack = [value];
  let length = 0;
  while (stack.length) {
    const item = stack.pop();
    if (typeof item === 'string') length += item.length;
    else if (Array.isArray(item)) {
      for (const child of item) stack.push(child);
    } else if (item !== null && typeof item === 'object')
      for (const [key, child] of Object.entries(item)) {
        length += key.length;
        stack.push(child);
      }
    if (length > limits.maxBytes) return { ok: false, reason: 'byte-limit' };
  }
  const raw = JSON.stringify(value);
  if (raw.length > limits.maxBytes || new TextEncoder().encode(raw).byteLength > limits.maxBytes)
    return { ok: false, reason: 'byte-limit' };
  return { ok: true, value: JSON.parse(raw) as JsonValue };
}
