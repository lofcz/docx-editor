import { independentPropertyChild } from './property-boundaries';
import {
  WML_NAMESPACE_URI,
  parentNodeOf,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import { readReviewGroups } from './groups';
import { currentRevisionPropertyContent } from './revision-property-content';
import { propertyGroupMembersKey, type PropertyGroupScope } from './property-history-contract';

/** Read complete source and proposed property identity, without allocating native operations. */
export function freshPropertyGroupScopes(
  before: OoxmlPackage,
  after: OoxmlPackage,
  freshIds: ReadonlySet<string>
): readonly PropertyGroupScope[] {
  const read = readReviewGroups(after);
  if (!read.success) throw new Error(read.error);
  const index = (root: OoxmlNode) => {
    const result = new Map<string, OoxmlNode>(),
      pending = [root];
    while (pending.length) {
      const node = pending.pop()!;
      result.set(node.id, node);
      if (node.kind !== 'textValue') pending.push(...node.children);
    }
    return result;
  };
  const sources = new Map([...before.parts].map(([name, part]) => [name, index(part.root)]));
  const targets = new Map([...after.parts].map(([name, part]) => [name, index(part.root)]));
  const output: PropertyGroupScope[] = [];
  for (const group of read.groups) {
    if (
      !freshIds.has(group.id) ||
      group.properties.length ||
      !group.native.length ||
      group.readOnly
    )
      continue;
    const ids = new Set<string>();
    const anchors: { parentId: string; childName: string }[] = [];
    const add = (node: OoxmlNode) => {
      const pending = [node];
      while (pending.length) {
        const item = pending.pop()!;
        ids.add(item.id);
        if (item.kind !== 'textValue')
          pending.push(
            ...item.children.filter(
              (child) =>
                child.kind === 'textValue' ||
                item.namespaceUri !== WML_NAMESPACE_URI ||
                child.namespaceUri !== WML_NAMESPACE_URI ||
                !independentPropertyChild(item.localName, child.localName)
            )
          );
      }
    };
    let complete = true;
    for (const member of group.native) {
      const part = after.parts.get(member.partName);
      for (const id of member.siteNodeIds) {
        const site = targets.get(member.partName)?.get(id);
        if (!part || !site || !currentRevisionPropertyContent(part, site)) {
          complete = false;
          break;
        }
        const parent = parentNodeOf(part, id);
        const original = parent && sources.get(member.partName)?.get(parent.id);
        if (!parent) {
          complete = false;
          break;
        }
        if (!original) {
          const ancestor = parentNodeOf(part, parent.id);
          // Numbering constructors have complete conditional ancestry support. Other
          // new property holders still need their own constructor dependency rules.
          if (
            parent.namespaceUri !== WML_NAMESPACE_URI ||
            parent.localName !== 'pPr' ||
            parent.children.some(
              (child) =>
                child.kind === 'textValue' ||
                child.namespaceUri !== WML_NAMESPACE_URI ||
                !['numPr', 'pPrChange'].includes(child.localName)
            ) ||
            !ancestor ||
            !sources.get(member.partName)?.has(ancestor.id)
          ) {
            complete = false;
            break;
          }
          anchors.push({ parentId: ancestor.id, childName: parent.localName });
        } else add(original);
        add(parent);
      }
      if (!complete) break;
    }
    if (complete)
      output.push(
        Object.freeze({
          groupId: group.id,
          membersKey: propertyGroupMembersKey(group),
          nodeIds: ids,
          ...(anchors.length ? { anchors: Object.freeze(anchors) } : {}),
        })
      );
  }
  const claimed = new Set<string>();
  for (const group of output)
    for (const id of group.nodeIds) {
      if (claimed.has(id))
        throw new Error('Independent property groups cannot share native ownership.');
      claimed.add(id);
    }
  return Object.freeze(output);
}
