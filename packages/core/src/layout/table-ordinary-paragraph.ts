import { WML_NAMESPACE_URI, type OoxmlNode } from '../store/package/ooxml-tree.ts';
import { MAX_XML_DEPTH } from '../store/package/xml-reader.ts';
import { MAX_PART_SCAN_ELEMENTS } from '../store/package/drawing-projection.ts';

const allowedElements = new Set([
  'adjustRightInd',
  'bCs',
  'cnfStyle',
  'p',
  'pPr',
  'r',
  'rPr',
  't',
  'tab',
  'br',
  'b',
  'i',
  'u',
  'sz',
  'szCs',
  'rFonts',
  'color',
  'lang',
  'spacing',
  'jc',
  'ind',
  'keepNext',
  'keepLines',
  'widowControl',
  'contextualSpacing',
  'pStyle',
  'rStyle',
  'noProof',
  'proofErr',
  'kern',
  'position',
  'vertAlign',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'shd',
  'highlight',
]);
const paragraphCaches = {
  ordinary: new WeakMap<OoxmlNode, boolean>(),
  drawing: new WeakMap<OoxmlNode, boolean>(),
  revisions: new WeakMap<OoxmlNode, boolean>(),
  probe: new WeakMap<OoxmlNode, boolean>(),
};
const controlMetadata = new Set(['alias', 'tag', 'id', 'lock']);
const numberingElements = new Set(['numPr', 'numId', 'ilvl']);

function ordinaryParagraph(root: OoxmlNode, policy: keyof typeof paragraphCaches): boolean {
  const cache = paragraphCaches[policy];
  const known = cache.get(root);
  if (known !== undefined) return known;
  let visited = 0;
  const visit = (node: OoxmlNode, depth: number, parent = ''): boolean => {
    if (++visited > MAX_PART_SCAN_ELEMENTS || depth > MAX_XML_DEPTH) return false;
    if (node.kind === 'textValue') return true;
    if (node.namespaceUri !== WML_NAMESPACE_URI) return false;
    if (policy === 'probe' && node.kind === 'contentControl' && parent === 'p') {
      const [props, content] = node.children;
      if (
        node.children.length !== 2 ||
        props?.kind !== 'contentControlProperties' ||
        content?.kind !== 'contentControlContent'
      )
        return false;
      visited += 2 + props.children.length;
      if (visited > MAX_PART_SCAN_ELEMENTS) return false;
      if (
        !props.children.every(
          (c: OoxmlNode) =>
            c.kind !== 'textValue' &&
            c.namespaceUri === WML_NAMESPACE_URI &&
            controlMetadata.has(c.localName) &&
            c.children.length === 0
        )
      )
        return false;
      return content.children.every((c) => c.kind === 'run' && visit(c, depth + 2, 'sdtContent'));
    }
    // A saved pagination marker carries no content or break instruction.
    if (node.localName === 'lastRenderedPageBreak') return node.children.length === 0;
    if (
      !allowedElements.has(node.localName) &&
      !(policy === 'drawing' && numberingElements.has(node.localName)) &&
      !(
        policy === 'revisions' &&
        ((['ins', 'del'].includes(node.localName) && ['p', 'ins', 'del'].includes(parent)) ||
          node.localName === 'delText')
      )
    )
      return false;
    return node.children.every((child) => visit(child, depth + 1, node.localName));
  };
  const result = visit(root, 0);
  cache.set(root, result);
  return result;
}

export function ordinaryTableParagraph(root: OoxmlNode): boolean {
  return ordinaryParagraph(root, 'ordinary');
}

/** Numbering affects paragraph flow, but does not introduce projected drawing atoms. */
export function ordinaryDrawingParagraph(root: OoxmlNode): boolean {
  return ordinaryParagraph(root, 'drawing');
}

/** Ordinary cell content with inline insertion and deletion wrappers. */
export function editableTableParagraph(root: OoxmlNode): boolean {
  return ordinaryParagraph(root, 'revisions');
}

/** Plain inline content controls can share a row measurement within one layout pass. */
export function ordinaryRowProbeParagraph(root: OoxmlNode): boolean {
  return ordinaryParagraph(root, 'probe');
}
