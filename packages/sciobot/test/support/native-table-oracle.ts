import { findNode, type OoxmlNode, type OoxmlPart } from '@docx-editor.dev/core/store';

export const nodeChildren = (node: OoxmlNode): readonly OoxmlNode[] =>
  node.kind === 'textValue' ? [] : node.children;
export const namedChild = (node: OoxmlNode, name: string) =>
  nodeChildren(node).find((entry) => entry.kind !== 'textValue' && entry.localName === name);
export const tableRows = (part: OoxmlPart, id: string) =>
  nodeChildren(findNode(part, id)!).filter((node) => node.kind === 'tableRow');
export const rowCells = (row: OoxmlNode) =>
  nodeChildren(row).filter((node) => node.kind === 'tableCell');
