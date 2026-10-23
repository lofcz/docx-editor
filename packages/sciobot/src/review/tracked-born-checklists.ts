import {
  contentControlPropertiesOf,
  createNodeIdAllocator,
  findNode,
  moveNode,
  replaceChildren,
  withPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlPackage,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import { refuseWrite } from '../diff/fragment-seeds';
import { SCIOBOT_CHECKLIST_TAG } from '../diff/lists';

const childElement = (node: OoxmlNode | undefined, name: string): OoxmlNode | undefined =>
  node && node.kind !== 'textValue'
    ? node.children.find(
        (child) =>
          child.kind !== 'textValue' &&
          child.namespaceUri === WML_NAMESPACE_URI &&
          child.localName === name
      )
    : undefined;

/** New controls need content insertion ownership as well as the paragraph's inserted mark. */
export function withBornChecklistOwnership(
  pkg: OoxmlPackage,
  partName: string,
  paragraphIds: readonly string[]
): OoxmlPackage {
  let working = pkg;
  const mint = createNodeIdAllocator(pkg.parts.get(partName)!);
  for (const id of paragraphIds) {
    const part = working.parts.get(partName)!;
    const paragraph = findNode(part, id);
    if (paragraph?.kind !== 'paragraph')
      refuseWrite('The born checklist paragraph disappeared.', 'apply-failed');
    const controls = paragraph.children.filter(
      (child) =>
        child.kind === 'contentControl' &&
        contentControlPropertiesOf(child).tag === SCIOBOT_CHECKLIST_TAG
    );
    if (!controls.length) continue;
    const mark = childElement(childElement(childElement(paragraph, 'pPr'), 'rPr'), 'ins');
    if (!mark || mark.kind === 'textValue')
      refuseWrite('A born checklist needs its insertion revision.', 'apply-failed');
    const address = (node: OoxmlElement) =>
      JSON.stringify(
        ['id', 'author', 'date'].map(
          (name) =>
            node.attributes.find(
              (attribute) =>
                attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
            )?.value ?? null
        )
      );
    const insertion = paragraph.children.find(
      (child) => child.kind === 'revisionInsert' && address(child) === address(mark)
    );
    if (insertion) {
      // Match the provider's normal form before content proofs are captured. Adjacent
      // equal insertions coalesce on publication, so put the control in the existing one.
      let movedPart = part;
      for (const [index, control] of controls.entries()) {
        const moved = moveNode(movedPart, control.id, insertion.id, index);
        if (!moved.ok)
          refuseWrite('The born checkbox could not join its insertion.', 'apply-failed');
        movedPart = moved.part;
      }
      working = withPart(working, movedPart);
      continue;
    }
    const children = paragraph.children.map((child) =>
      controls.includes(child)
        ? ({
            kind: 'revisionInsert',
            id: mint(),
            namespaceUri: WML_NAMESPACE_URI,
            localName: 'ins',
            prefix: 'w',
            namespaceBindings: [],
            attributes: mark.attributes.filter(
              (attribute) =>
                attribute.namespaceUri === WML_NAMESPACE_URI &&
                ['id', 'author', 'date'].includes(attribute.localName)
            ),
            children: [child],
          } as OoxmlElement)
        : child
    );
    const replaced = replaceChildren(part, id, children);
    if (!replaced.ok)
      refuseWrite('The born checklist insertion failed native validation.', 'apply-failed');
    working = withPart(working, replaced.part);
  }
  return working;
}
