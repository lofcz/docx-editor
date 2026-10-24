import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  collectRevisionSites,
  parentNodeOf,
  readOoxmlPackage,
  readOoxmlPart,
  withPart,
  type OoxmlNode,
  type OoxmlElement,
} from '@docx-editor.dev/core/store';
import { docxFixture, W } from '../../test/support/document';
import { persistReviewGroups, readReviewGroups } from './groups';
import { withInitializedReviewGroupStore } from './group-record-store';
import { freshPropertyGroupScopes } from './property-group-scopes';
import { propertyBoundaryBody, propertyBoundaryOps } from '../../test/support/property-boundaries';
test('paragraph, mark-format, and section ownership stay independent in one property publication', () => {
  const loaded = readOoxmlPackage(docxFixture(propertyBoundaryBody));
  if (!loaded.ok) throw new Error(loaded.reason);
  const prepared = withInitializedReviewGroupStore(loaded.package);
  if (!prepared.ok) throw new Error(prepared.reason);
  const before = prepared.pkg;
  const part = before.parts.get(before.mainDocumentPart)!;
  const nodes: OoxmlNode[] = [];
  const walk = (node: OoxmlNode) => {
    nodes.push(node);
    if (node.kind !== 'textValue') node.children.forEach(walk);
  };
  walk(part.root);
  const paragraph = nodes.find((node) => node.kind === 'paragraph')!;
  let after = part;
  for (const op of propertyBoundaryOps(paragraph.id)) {
    const result = applyTreeOp(after, op);
    if (!result.ok) throw new Error(result.reason);
    after = result.part;
  }
  const pkg = persistReviewGroups(before, withPart(before, after), [], false);
  const groups = readReviewGroups(pkg);
  if (!groups.success) throw new Error(groups.error);
  const fresh = groups.groups.filter((group) =>
    group.native.every((member) => member.address.author !== 'Peer')
  );
  expect(fresh).toHaveLength(2);
  const scopes = freshPropertyGroupScopes(before, pkg, new Set(fresh.map((group) => group.id)));
  expect(scopes).toHaveLength(2);
  const marker = nodes.find(
    (node) => node.kind !== 'textValue' && node.namespaceUri === W && node.localName === 'ins'
  )!;
  const section = nodes.find(
    (node) => node.kind !== 'textValue' && node.namespaceUri === W && node.localName === 'sectPr'
  )!;
  const mark = parentNodeOf(part, marker.id)!;
  const foreign = nodes.find(
    (node) =>
      node.kind !== 'textValue' && node.namespaceUri === 'urn:opaque' && node.localName === 'rPr'
  )!;
  const paragraphGroup = fresh.find((group) => group.native[0]!.address.author === 'Alice')!;
  const markGroup = fresh.find((group) => group.native[0]!.address.author === 'Bot')!;
  const pScope = scopes.find((scope) => scope.groupId === paragraphGroup.id)!;
  const rScope = scopes.find((scope) => scope.groupId === markGroup.id)!;
  expect(pScope.nodeIds.has(foreign.id)).toBe(true);
  expect(pScope.nodeIds.has(mark.id)).toBe(false);
  expect(pScope.nodeIds.has(section.id)).toBe(false);
  expect(pScope.nodeIds.has(marker.id)).toBe(false);
  expect(rScope.nodeIds.has(mark.id)).toBe(true);
  expect(rScope.nodeIds.has(marker.id)).toBe(false);
  expect([...pScope.nodeIds].filter((id) => rScope.nodeIds.has(id))).toEqual([]);
  expect(collectRevisionSites(pkg.parts.get(pkg.mainDocumentPart)!)).toHaveLength(3);
});

test('row and cell property scopes exclude independent insertion and merge markers', () => {
  const marker = (name: string, id: number) => `<w:${name} w:id="${id}" w:author="Peer"/>`;
  const body =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="200"/>' +
    marker('ins', 71) +
    '</w:trPr><w:tc><w:tcPr><w:shd w:fill="FFFFFF"/>' +
    marker('cellIns', 72) +
    marker('cellMerge', 73) +
    '</w:tcPr><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>';
  const loaded = readOoxmlPackage(docxFixture(body));
  if (!loaded.ok) throw new Error(loaded.reason);
  const initialized = withInitializedReviewGroupStore(loaded.package);
  if (!initialized.ok) throw new Error(initialized.reason);
  const before = initialized.pkg;
  let part = before.parts.get(before.mainDocumentPart)!;
  const nodes: OoxmlNode[] = [];
  const walk = (node: OoxmlNode) => {
    nodes.push(node);
    if (node.kind !== 'textValue') node.children.forEach(walk);
  };
  walk(part.root);
  // Imported pending property groups can coexist with independent structural revisions.
  const addRecord = (node: OoxmlNode): OoxmlNode => {
    if (node.kind === 'textValue') return node;
    const container = node.localName;
    if (node.namespaceUri === W && ['trPr', 'tcPr'].includes(container)) {
      const author = container === 'trPr' ? 'Alice' : 'Bot';
      const record = readOoxmlPart(
        `<w:${container}Change xmlns:w="${W}" w:id="${container === 'trPr' ? 81 : 82}" w:author="${author}"><w:${container}/></w:${container}Change>`,
        { name: '/record.xml', contentType: 'application/xml' }
      );
      if (!record.ok) throw new Error(record.reason);
      const unique = (item: OoxmlNode): OoxmlNode =>
        item.kind === 'textValue'
          ? { ...item, id: `${container}:record:${item.id}` }
          : ({
              ...item,
              id: `${container}:record:${item.id}`,
              children: item.children.map(unique),
            } as OoxmlElement);
      return { ...node, children: [...node.children, unique(record.part.root)] } as OoxmlElement;
    }
    return { ...node, children: node.children.map(addRecord) } as OoxmlElement;
  };
  part = { ...part, root: addRecord(part.root) as typeof part.root };
  const pkg = persistReviewGroups(before, withPart(before, part), [], false);
  const groups = readReviewGroups(pkg);
  if (!groups.success) throw new Error(groups.error);
  const fresh = groups.groups.filter((group) =>
    group.native.every((member) => member.address.author !== 'Peer')
  );
  expect(fresh).toHaveLength(1);
  expect(fresh[0]!.native).toHaveLength(2);
  const scopes = freshPropertyGroupScopes(before, pkg, new Set(fresh.map((group) => group.id)));
  expect(scopes).toHaveLength(1);
  for (const scope of scopes) {
    for (const node of nodes.filter(
      (node) =>
        node.kind !== 'textValue' && ['ins', 'cellIns', 'cellMerge'].includes(node.localName)
    ))
      expect(scope.nodeIds.has(node.id)).toBe(false);
  }
  for (const node of nodes.filter(
    (node) => node.kind !== 'textValue' && ['trPr', 'tcPr'].includes(node.localName)
  ))
    expect(scopes[0]!.nodeIds.has(node.id)).toBe(true);
});
