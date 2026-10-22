import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';

export const INSERTED_PARAGRAPH_PROPERTY_NS =
  'urn:docx-editor:inserted-paragraph-property-presence:1';

/** Retain only source container presence, never formatting or document text. */
export function withInsertedParagraphPropertyPresence(
  mark: OoxmlElement,
  properties: OoxmlNode | undefined,
  runProperties: OoxmlNode | undefined
): OoxmlElement {
  if (!properties) return mark;
  const prefix = `paragraphOrigin${crypto.randomUUID().replace(/-/g, '')}`;
  return {
    ...mark,
    namespaceBindings: [
      ...mark.namespaceBindings,
      { prefix, namespaceUri: INSERTED_PARAGRAPH_PROPERTY_NS },
    ],
    attributes: [
      ...mark.attributes,
      {
        kind: 'genericExtension',
        namespaceUri: INSERTED_PARAGRAPH_PROPERTY_NS,
        prefix,
        localName: 'properties',
        value: runProperties ? 'pPr rPr' : 'pPr',
      },
    ],
  } as OoxmlElement;
}

/** A copied source container survives the inserted mark's acceptance after save/reopen. */
export function preservesInsertedParagraphProperty(node: OoxmlNode, kind: 'pPr' | 'rPr'): boolean {
  if (node.kind === 'textValue') return false;
  const fields = node.attributes.filter(
    (attr) => attr.namespaceUri === INSERTED_PARAGRAPH_PROPERTY_NS
  );
  return (
    fields.length === 1 &&
    fields[0]!.localName === 'properties' &&
    (fields[0]!.value === 'pPr rPr' || (kind === 'pPr' && fields[0]!.value === 'pPr'))
  );
}
