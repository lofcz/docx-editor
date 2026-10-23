import { secureRandomUuid } from '../package/random-uuid.ts';
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
  return withPropertyPresence(mark, runProperties ? 'pPr rPr' : 'pPr');
}

/** A structural insertion keeps the holders its direct native command already created. */
export function withInsertedTablePropertyPresence(
  mark: OoxmlElement,
  kind: 'trPr' | 'tcPr'
): OoxmlElement {
  return withPropertyPresence(mark, kind);
}

function withPropertyPresence(
  mark: OoxmlElement,
  properties: 'pPr' | 'pPr rPr' | 'trPr' | 'tcPr' | 'absent trPr' | 'absent tcPr'
): OoxmlElement {
  const prefix = `paragraphOrigin${secureRandomUuid().replace(/-/g, '')}`;
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
        value: properties,
      },
    ],
  } as OoxmlElement;
}

/** A copied source container survives the inserted mark's acceptance after save/reopen. */
export function preservesInsertedParagraphProperty(node: OoxmlNode, kind: 'pPr' | 'rPr'): boolean {
  return preservesInsertedPropertyPresence(node, kind);
}

export function preservesInsertedPropertyPresence(
  node: OoxmlNode,
  kind: 'pPr' | 'rPr' | 'trPr' | 'tcPr'
): boolean {
  const value = propertyPresence(node);
  return value !== null && !value.startsWith('absent ') && value.split(' ').includes(kind);
}

export function withAbsentTablePropertyPresence(
  mark: OoxmlElement,
  kind: 'trPr' | 'tcPr'
): OoxmlElement {
  return withPropertyPresence(mark, `absent ${kind}`);
}

export function omitsInsertedTableProperty(node: OoxmlNode, kind: 'trPr' | 'tcPr'): boolean {
  return propertyPresence(node) === `absent ${kind}`;
}

export function hasInsertedPropertyPresence(node: OoxmlNode): boolean {
  return propertyPresence(node) !== null;
}

function propertyPresence(node: OoxmlNode): string | null {
  if (node.kind === 'textValue') return null;
  const fields = node.attributes.filter(
    (attr) => attr.namespaceUri === INSERTED_PARAGRAPH_PROPERTY_NS
  );
  return fields.length === 1 &&
    fields[0]!.localName === 'properties' &&
    ['pPr', 'pPr rPr', 'trPr', 'tcPr', 'absent trPr', 'absent tcPr'].includes(fields[0]!.value)
    ? fields[0]!.value
    : null;
}
