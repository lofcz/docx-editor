import {
  collectRevisionSites,
  findNode,
  replaceChildren,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';

/**
 * A guarded post-operation edit. Leading insertion keeps an unseen append after the
 * replacement in either Yjs client order. Move only new sibling wrappers owned by this
 * transaction; existing revisions and inline container ownership remain intact.
 */
export function leadingReplacementEdit(
  before: OoxmlPart,
  paragraphIds: ReadonlySet<string>,
  author: string,
  date: string
): (pkg: OoxmlPackage) => OoxmlPackage {
  const prior = new Set(collectRevisionSites(before).map((site) => site.node.id));
  return (pkg) => {
    let part = pkg.parts.get(before.name);
    if (!part) throw new Error('The replacement story disappeared.');
    const fresh = new Set(
      collectRevisionSites(part)
        .filter((site) => {
          if (prior.has(site.node.id) || site.refused || site.paragraphMark || site.propertyChange)
            return false;
          const value = (name: string) =>
            site.node.attributes.find(
              (attribute) =>
                attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
            )?.value;
          return value('author') === author && value('date') === date;
        })
        .map((site) => site.node.id)
    );
    if (!fresh.size) return pkg;
    const matches = (node: OoxmlNode | undefined, kind: 'revisionInsert' | 'revisionDelete') =>
      node?.kind === kind && fresh.has(node.id);
    const visit = (node: OoxmlNode) => {
      if (node.kind === 'textValue') return;
      for (const child of node.children) visit(child);
      const current = findNode(part!, node.id);
      if (!current || current.kind === 'textValue')
        throw new Error('A replacement container disappeared.');
      const children: OoxmlNode[] = [];
      let changed = false;
      for (let at = 0; at < current.children.length; ) {
        if (!matches(current.children[at], 'revisionDelete')) {
          children.push(current.children[at++]!);
          continue;
        }
        const start = at;
        while (matches(current.children[at], 'revisionDelete')) at++;
        const middle = at;
        while (matches(current.children[at], 'revisionInsert')) at++;
        if (at > middle) {
          children.push(
            ...current.children.slice(middle, at),
            ...current.children.slice(start, middle)
          );
          changed = true;
        } else children.push(...current.children.slice(start, middle));
      }
      if (changed) {
        const result = replaceChildren(part!, current.id, children);
        if (!result.ok) throw new Error('Leading replacement ordering failed native validation.');
        part = result.part;
      }
    };
    for (const id of paragraphIds) {
      const paragraph = findNode(part, id);
      if (paragraph?.kind !== 'paragraph') throw new Error('A replacement paragraph disappeared.');
      visit(paragraph);
    }
    return part === pkg.parts.get(before.name) ? pkg : withPart(pkg, part);
  };
}
