/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { OoxmlNode, OoxmlPart } from '@docx-editor.dev/core/store';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const O = 'urn:schemas-microsoft-com:office:office';
interface Summary {
  readonly ids: readonly string[];
  readonly own: readonly string[];
  readonly children: readonly Summary[];
}
const EMPTY: Summary = Object.freeze({ ids: [], own: [], children: [] });
export type ImageReferenceState = readonly (readonly [string, Summary])[];
/** Changed text yields the same reference summary; unchanged native subtrees cost one lookup. */
export class ImageReferenceIndex {
  private readonly nodes = new WeakMap<OoxmlNode, Summary>();
  private readonly previous = new Map<string, Summary>();
  private state: ImageReferenceState = [];
  private summary(node: OoxmlNode): Summary {
    if (node.kind === 'textValue') return EMPTY;
    const held = this.nodes.get(node);
    if (held) return held;
    const own = node.attributes
      .filter((a) => a.namespaceUri === R || (a.namespaceUri === O && a.localName === 'relid'))
      .map((a) => a.value);
    const children: Summary[] = [];
    for (const child of node.children) {
      const refs = this.summary(child);
      if (refs !== EMPTY) children.push(refs);
    }
    if (!own.length && !children.length) {
      this.nodes.set(node, EMPTY);
      return EMPTY;
    }
    const old = this.previous.get(node.id);
    if (
      old &&
      own.length === old.own.length &&
      own.every((v, i) => v === old.own[i]) &&
      children.length === old.children.length &&
      children.every((v, i) => v === old.children[i])
    ) {
      this.nodes.set(node, old);
      return old;
    }
    const next: Summary = {
      own,
      children,
      ids: [...new Set([...own, ...children.flatMap((child) => child.ids)])],
    };
    this.previous.set(node.id, next);
    this.nodes.set(node, next);
    return next;
  }
  read(parts: ReadonlyMap<string, OoxmlPart>): ImageReferenceState {
    const next: (readonly [string, Summary])[] = [];
    for (const [owner, part] of parts) {
      const summary = this.summary(part.root);
      if (summary !== EMPTY) next.push([owner, summary]);
    }
    if (
      next.length === this.state.length &&
      next.every(
        ([owner, summary], i) => owner === this.state[i]![0] && summary === this.state[i]![1]
      )
    )
      return this.state;
    return (this.state = next);
  }
}
