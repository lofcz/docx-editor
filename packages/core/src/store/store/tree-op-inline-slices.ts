import type { OoxmlNode, OoxmlRunNode } from '../package/ooxml-tree.ts';
import { textElement } from './tree-op-inline-elements.ts';
import {
  recordNativeTextSlice,
  recordNativeTextMerge,
} from '../package/canonical-primitive-capture.ts';
const XML = 'http://www.w3.org/XML/1998/namespace';
/** Keep source element metadata on native range slices; mint only their structural identities. */
export function inlineTextSlice(
  mint: () => string,
  value: string,
  source: OoxmlNode,
  sourceOffset: number,
  retainIdentity = false
): OoxmlNode {
  if (source.kind !== 'text' && source.kind !== 'deletedText') return textElement(mint, value);
  const created = textElement(mint, value, source.kind);
  // An internal source space becomes a boundary space after a split. Word must preserve it.
  const attributes = /^\s|\s$/.test(value)
    ? source.attributes.filter(
        (attr) =>
          !(attr.namespaceUri === XML && attr.localName === 'space' && attr.value === 'default')
      )
    : source.attributes;
  // One trailing fragment keeps the observed leaf. Other fragments mint independent IDs.
  // Shared text can then anchor unchanged source characters across an inserted control.
  const retained =
    retainIdentity && source.children.length === 1 && source.children[0]?.kind === 'textValue'
      ? { ...source.children[0], value }
      : null;
  const leaf = retained ?? (created.kind !== 'textValue' ? created.children[0] : undefined);
  if (
    source.children.length === 1 &&
    source.children[0]?.kind === 'textValue' &&
    leaf?.kind === 'textValue'
  )
    recordNativeTextSlice(source.children[0], leaf, sourceOffset);
  return {
    ...source,
    id: retained ? source.id : created.id,
    attributes,
    children: retained ? [retained] : created.kind === 'textValue' ? [] : created.children,
  } as OoxmlNode;
}
/** Join adjacent text leaves without treating the result as one source slice. */
export function inlineTextMerged(mint: () => string, left: OoxmlNode, right: OoxmlNode): OoxmlNode {
  if (left.kind === 'textValue' || right.kind === 'textValue')
    throw new Error('Expected text elements.');
  const a = left.children[0];
  const b = right.children[0];
  if (a?.kind !== 'textValue' || b?.kind !== 'textValue') throw new Error('Expected text leaves.');
  const created = textElement(
    mint,
    a.value + b.value,
    left.kind === 'deletedText' ? 'deletedText' : 'text'
  );
  if (created.kind === 'textValue' || created.children[0]?.kind !== 'textValue')
    throw new Error('Expected text leaf.');
  recordNativeTextMerge(a, b, created.children[0]);
  return { ...left, id: created.id, children: created.children } as OoxmlNode;
}

export function inlineRunSlice(
  mint: () => string,
  source: OoxmlRunNode,
  children: readonly OoxmlNode[]
): OoxmlNode {
  return { ...source, id: mint(), children } as OoxmlNode;
}

/** Source text metadata, including local QName bindings, must agree before text coalesces. */
export function sameInlineTextMetadata(left: OoxmlNode, right: OoxmlNode): boolean {
  if (left.kind === 'textValue' || right.kind === 'textValue') return false;
  if (
    !left.attributes.length &&
    !right.attributes.length &&
    !left.namespaceBindings.length &&
    !right.namespaceBindings.length
  )
    return true;
  const attributes = (node: typeof left) =>
    JSON.stringify(
      node.attributes.map((attr) => [attr.namespaceUri, attr.localName, attr.value]).sort()
    );
  const bindings = (node: typeof left) =>
    JSON.stringify(
      node.namespaceBindings.map((binding) => [binding.prefix, binding.namespaceUri]).sort()
    );
  return attributes(left) === attributes(right) && bindings(left) === bindings(right);
}
