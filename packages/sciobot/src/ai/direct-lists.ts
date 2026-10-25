import { copyListInstanceTemplate, setListInstanceStart } from './list-instance-template';
import { writeDormantListState } from '../diff/list-metadata';
import { styleListInstance } from './direct-list-style';
import {
  applyTreeOp,
  createNodeIdAllocator,
  contentControlPropertiesOf,
  directParagraphProperties,
  ensureListDefinition,
  findNode,
  insertChildren,
  readOoxmlPart,
  withPart,
  type OoxmlPackage,
  type OoxmlNode,
  type OoxmlPart,
  type TreeDocOp,
} from '@docx-editor.dev/core/store';
import { SCIOBOT_CHECKLIST_TAG } from '../diff/lists';
import type { ListSourceState, ListTarget } from '../diff/list-types';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** One paragraph's list destination inside a direct transaction. */
export interface DirectListAssignment {
  readonly paragraphId: string;
  /** Null for a paragraph created in this transaction. */
  readonly before: ListSourceState | null;
  readonly target: ListTarget;
}

/** Allocation state belongs to one plan execution, never its preview or a later plan. */
export interface DirectListContext {
  readonly instances: Map<string, string>;
  readonly levels: Set<string>;
}
export function createDirectListContext(): DirectListContext {
  return { instances: new Map(), levels: new Set() };
}

class ListWriteError extends Error {}

/** Controlled XML seeds need fresh identities in their destination, including across calls. */
function freshSeeds(part: OoxmlPart, nodes: readonly OoxmlNode[]): OoxmlNode[] {
  const mint = createNodeIdAllocator(part);
  const copy = (node: OoxmlNode): OoxmlNode =>
    node.kind === 'textValue'
      ? { ...node, id: mint() }
      : ({ ...node, id: mint(), children: node.children.map(copy) } as OoxmlNode);
  return nodes.map(copy);
}

function seedNodes(part: OoxmlPart, xml: string, name: string) {
  const parsed = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${xml}</w:body></w:document>`,
    { name, contentType: part.contentType }
  );
  if (!parsed.ok) throw new ListWriteError(`Invalid list seed: ${parsed.reason}`);
  const body = parsed.part.root.children.find(
    (node) => node.kind !== 'textValue' && node.localName === 'body'
  );
  return body && body.kind !== 'textValue' ? freshSeeds(part, body.children) : [];
}

/** Remove one source indentation step while keeping trailing and opaque geometry. */
export function removedListProperties(part: OoxmlPart, paragraphId: string, steps: number) {
  const properties = directParagraphProperties(part, paragraphId),
    attributes = { ...properties.find((property) => property.localName === 'ind')?.attributes };
  for (const name of [
    'left',
    'start',
    'hanging',
    'firstLine',
    'firstLineChars',
    'hangingChars',
    'leftChars',
    'startChars',
  ])
    delete attributes[name];
  if (steps > 0) attributes.left = String(steps * 360);
  return [
    ...properties.filter((property) => property.localName !== 'ind'),
    ...(Object.keys(attributes).length ||
    properties.some((property) => property.localName === 'ind')
      ? [{ localName: 'ind', attributes }]
      : []),
  ];
}

/**
 * Apply owned list membership after native identity allocation. New instances share one Word
 * definition per plan key; existing instances keep their definitions and later ordinals.
 */
export function applyDirectListAssignments(
  input: OoxmlPackage,
  partName: string,
  assignments: readonly DirectListAssignment[],
  context: DirectListContext = createDirectListContext(),
  revision?: Extract<TreeDocOp, { op: 'setListNumbering' }>['revision']
): OoxmlPackage {
  if (
    revision &&
    assignments.some(({ before, target }) => before?.checkbox || target.kind === 'checklist')
  )
    throw new ListWriteError('Tracked checkbox changes need the checklist executor.');
  let pkg = input,
    seeds = 0;
  const { instances, levels } = context;
  const part = () => pkg.parts.get(partName)!;
  const apply = (op: TreeDocOp) => {
    const result = applyTreeOp(part(), op);
    if (!result.ok) throw new ListWriteError(`Native list write refused: ${result.reason}.`);
    pkg = withPart(pkg, result.part);
  };
  const numId = (target: Extract<ListTarget, { kind: 'ordered' | 'bullet' }>): string => {
    if (target.instance.kind === 'existing') return target.instance.numId;
    let id = instances.get(target.instance.key);
    if (!id) {
      const made = ensureListDefinition(pkg, target.kind);
      if (!made) throw new ListWriteError('The native engine refused a new list definition.');
      pkg = target.instance.templateNumId
        ? copyListInstanceTemplate(made.pkg, target.instance.templateNumId, made.numId)
        : target.style
          ? styleListInstance(made.pkg, made.numId, target.style)
          : made.pkg;
      id = made.numId;
      instances.set(target.instance.key, id);
    }
    const ancestorKey = `${id}:ancestors`;
    if (target.instance.ancestorStarts && !levels.has(ancestorKey)) {
      for (const [level, start] of target.instance.ancestorStarts.entries())
        pkg = setListInstanceStart(pkg, id, level, start);
      levels.add(ancestorKey);
    }
    const levelKey = `${id}:${target.level}`;
    if (target.instance.start !== undefined && !levels.has(levelKey)) {
      pkg = setListInstanceStart(pkg, id, target.level, target.instance.start);
      levels.add(levelKey);
    }
    return id;
  };
  for (const { paragraphId, before, target } of assignments) {
    const paragraph = findNode(part(), paragraphId);
    if (paragraph?.kind !== 'paragraph')
      throw new ListWriteError('A list paragraph disappeared during the transaction.');
    const checkbox = before?.checkbox ?? null;
    if (checkbox && target.kind !== 'checklist')
      apply({ op: 'removeContentControl', controlId: checkbox.controlId, keepContent: false });
    if (target.kind === 'checklist') {
      if (checkbox) {
        if (checkbox.checked !== target.checked)
          apply({
            op: 'setContentControlValue',
            controlId: checkbox.controlId,
            value: { kind: 'checkbox', checked: target.checked },
          });
      } else {
        const marker = seedNodes(
          part(),
          `<w:p><w:sdt xmlns:w14="${W14}"><w:sdtPr><w:tag w:val="${SCIOBOT_CHECKLIST_TAG}"/><w14:checkbox><w14:checked w14:val="${target.checked ? 1 : 0}"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r><w:t>${target.checked ? '☒' : '☐'}</w:t></w:r></w:sdtContent></w:sdt></w:p>`,
          `/sciobot-direct-list-${++seeds}.xml`
        )[0];
        const control =
          marker?.kind === 'paragraph' &&
          marker.children.find((node) => node.kind === 'contentControl');
        if (!control || !contentControlPropertiesOf(control).checkbox)
          throw new ListWriteError('Invalid native checklist marker.');
        const current = findNode(part(), paragraphId)!;
        if (current.kind !== 'paragraph') throw new ListWriteError('Missing list paragraph.');
        const inserted = insertChildren(
          part(),
          current.id,
          current.children[0]?.localName === 'pPr' ? 1 : 0,
          [control]
        );
        if (!inserted.ok) throw new ListWriteError('The native checklist marker was refused.');
        pkg = withPart(pkg, inserted.part);
      }
      if (!checkbox || before?.numbering)
        apply({ op: 'setListNumbering', paragraphId, numId: '0' });
    } else
      apply({
        op: 'setListNumbering',
        paragraphId,
        numId: target.kind === 'none' ? '0' : numId(target),
        ...(target.kind === 'none' ? {} : { level: target.level }),
        ...(revision ? { revision } : {}),
      });
    if (
      target.kind === 'checklist' ||
      checkbox ||
      (target.kind === 'none' && target.retainedIndent !== undefined)
    ) {
      const properties = directParagraphProperties(part(), paragraphId),
        attributes = { ...properties.find((property) => property.localName === 'ind')?.attributes };
      for (const name of ['left', 'start', 'hanging']) delete attributes[name];
      if (target.kind === 'checklist') {
        for (const name of [
          'firstLine',
          'firstLineChars',
          'hangingChars',
          'leftChars',
          'startChars',
        ])
          delete attributes[name];
        attributes.left = String(720 * (target.level + 1));
        attributes.hanging = '360';
      }
      apply({
        op: 'setParagraphProperties',
        paragraphId,
        ...(revision && target.kind === 'none' && target.retainedIndent !== undefined
          ? { revision }
          : {}),
        properties:
          target.kind === 'none' && target.retainedIndent !== undefined
            ? removedListProperties(part(), paragraphId, target.retainedIndent)
            : [
                ...properties.filter((property) => property.localName !== 'ind'),
                ...(Object.keys(attributes).length ? [{ localName: 'ind', attributes }] : []),
              ],
      });
    }
    const dormant =
      target.kind === 'none' || target.kind === 'checklist'
        ? undefined
        : (target.dormantChecked ?? before?.dormantChecked);
    pkg = withPart(pkg, writeDormantListState(part(), paragraphId, dormant));
  }
  return pkg;
}
