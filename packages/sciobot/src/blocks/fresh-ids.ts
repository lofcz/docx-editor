import {
  mintParaId,
  mintedParagraphIdentityAttributes,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';

/**
 * Fresh canonical IDs, plus a `w14:paraId` on every paragraph when the host uses them. Remote
 * publication replaces canonical IDs, and selections and peer carets address paragraphs by
 * `paraId`; without one, the caret placed here would name a paragraph that no longer exists.
 * The block ID seeds the paraIds, so two sessions inserting different blocks cannot collide.
 */
export function withFreshIds(
  node: OoxmlNode,
  next: () => string,
  identity: { readonly prefix: string | null; readonly used: Set<string>; readonly seed: string }
): OoxmlNode {
  if (node.kind === 'textValue') return { ...node, id: next() };
  let attributes = node.attributes;
  if (node.kind === 'paragraph' && identity.prefix !== null) {
    const minted = mintParaId(`${identity.seed}:${String(identity.used.size)}`, identity.used);
    identity.used.add(minted);
    attributes = [
      ...attributes,
      ...mintedParagraphIdentityAttributes(identity.prefix, minted),
    ] as typeof attributes;
  }
  return {
    ...node,
    id: next(),
    attributes,
    children: node.children.map((child) => withFreshIds(child, next, identity)),
  } as OoxmlNode;
}
