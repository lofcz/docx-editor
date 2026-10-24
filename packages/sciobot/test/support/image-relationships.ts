import { expect } from 'bun:test';
import {
  findNode,
  resolveImageRelationship,
  type OoxmlNode,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
interface ImageBlock {
  readonly controlId: string;
  readonly partName: string;
  readonly payload: {
    readonly type: string;
    readonly attributes: Readonly<Record<string, unknown>>;
  };
}
const embedded = (node: OoxmlNode): string[] =>
  node.kind === 'textValue'
    ? []
    : [
        ...node.attributes
          .filter(
            (attribute) =>
              attribute.namespaceUri ===
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships' &&
              attribute.localName === 'embed'
          )
          .map((attribute) => attribute.value),
        ...node.children.flatMap(embedded),
      ];
/** Check the actual native drawing relationship, not only its application provenance record. */
export function expectImageRelationships(pkg: OoxmlPackage, blocks: Iterable<ImageBlock>): void {
  for (const block of blocks) {
    if (block.payload.type !== 'inline_image') continue;
    const part = pkg.parts.get(block.partName)!;
    const control = findNode(part, block.controlId)!;
    const ids = embedded(control);
    expect(ids).toHaveLength(1);
    const related = resolveImageRelationship(pkg.relationships.get(part.name), part.name, ids[0]!);
    expect(related).toMatchObject({
      mode: 'internal',
      partName: block.payload.attributes.mediaPartName,
    });
  }
}
