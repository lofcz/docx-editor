import type { OoxmlContentControlNode, OoxmlNode } from './ooxml-tree.ts';
import { MC_NAMESPACE_URI } from './ooxml-shared.ts';

/** Durable ownership of application fallback content. @internal */
export const ATOMIC_CONTENT_NAMESPACE = 'urn:docx-editor.dev:atomic-content:1';

/** Only the supported marker opts a control into atomic fallback ownership. @internal */
export function isAtomicContentControl(node: OoxmlNode): boolean {
  return (
    node.kind === 'contentControl' &&
    node.attributes.some(
      (attribute) =>
        attribute.namespaceUri === ATOMIC_CONTENT_NAMESPACE &&
        attribute.localName === 'atomicContent' &&
        attribute.value === '1'
    )
  );
}

/** Mark the canonical control, preserving extension attributes through OOXML consumers. @internal */
export function withAtomicContentMarker(control: OoxmlContentControlNode): OoxmlContentControlNode {
  if (isAtomicContentControl(control)) return control;
  const used = new Set<string>();
  const pending: OoxmlNode[] = [control];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    if (node.prefix) used.add(node.prefix);
    for (const attribute of node.attributes) if (attribute.prefix) used.add(attribute.prefix);
    for (const binding of node.namespaceBindings) used.add(binding.prefix);
    for (const child of node.children) pending.push(child);
  }
  const fresh = (base: string) => {
    let prefix = base;
    for (let suffix = 2; used.has(prefix); suffix++) prefix = `${base}${suffix}`;
    used.add(prefix);
    return prefix;
  };
  const atomic = fresh('docxAtomic');
  const mc = fresh('mcAtomic');
  const attributes = control.attributes.filter(
    (attribute) =>
      !(
        attribute.namespaceUri === ATOMIC_CONTENT_NAMESPACE &&
        attribute.localName === 'atomicContent'
      )
  );
  const token = (name: string, value: string) => {
    const at = attributes.findIndex(
      (attribute) => attribute.namespaceUri === MC_NAMESPACE_URI && attribute.localName === name
    );
    const previous = at < 0 ? undefined : attributes[at];
    const values = new Set(previous?.value.trim().split(/\s+/).filter(Boolean) ?? []);
    values.add(value);
    const attribute = {
      kind: 'genericExtension' as const,
      namespaceUri: MC_NAMESPACE_URI,
      localName: name,
      prefix: mc,
      value: [...values].join(' '),
    };
    if (at < 0) attributes.push(attribute);
    else attributes[at] = attribute;
  };
  token('Ignorable', atomic);
  token('PreserveAttributes', `${atomic}:atomicContent`);
  attributes.push({
    kind: 'genericExtension',
    namespaceUri: ATOMIC_CONTENT_NAMESPACE,
    localName: 'atomicContent',
    prefix: atomic,
    value: '1',
  });
  return {
    ...control,
    attributes,
    namespaceBindings: [
      ...control.namespaceBindings,
      { prefix: atomic, namespaceUri: ATOMIC_CONTENT_NAMESPACE },
      { prefix: mc, namespaceUri: MC_NAMESPACE_URI },
    ],
  };
}
