import type { TargetAstNode } from '../projection/parse-target';
import type { NativeEditFailure } from '../contracts/document';

/** Text that XML can hold: no control characters other than tab, line feed, and return. */
export const xmlText = (value: unknown): value is string =>
  typeof value === 'string' &&
  !/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/.test(value) &&
  !/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value);

/** Recognized AST fields only: extension metadata must never disappear into native fallback. */
export function astFields(node: TargetAstNode, names: readonly string[]): boolean {
  return Object.keys(node).every(
    (key) =>
      key === 'position' ||
      names.includes(key) ||
      (key === 'data' &&
        node.data !== null &&
        typeof node.data === 'object' &&
        !Array.isArray(node.data) &&
        !Object.keys(node.data).length)
  );
}

/** A component hook refuses an MDX node with a stable failure code and message. */
export class MdxRefusal extends Error {
  constructor(
    message: string,
    readonly code: NativeEditFailure['code'] = 'invalid-input'
  ) {
    super(message);
    this.name = 'MdxRefusal';
  }
}
