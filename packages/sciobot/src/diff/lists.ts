import {
  buildNumberingIndex,
  buildStyleCascadeTable,
  cascadeParagraphFormatting,
  readNumPr,
  resolveNumberingLevel,
  resolveParagraphLayoutInputs,
  sha256FontBytes,
  withNumberingStyleLinks,
} from '@docx-editor.dev/core/layout';
import {
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  findNode,
  propertyContainer,
  readTwipsMeasure,
  resolveRelationship,
  stylesPartOf,
  type OoxmlNode,
  type OoxmlPackage,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem } from '../data/json';
import { isListTarget, type ListSourceState, type ListTarget } from './list-types';
import { createEditPlan, type CanonicalEdit, type EditPlanResult } from './plan';

export type { ListInstance, ListSourceState, ListTarget } from './list-types';
export const SCIOBOT_CHECKLIST_TAG = 'scio:list-checkbox';
export interface ListPlanSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly partName: string;
}
export interface ListMembershipPatch {
  readonly paragraphId: string;
  readonly target: ListTarget;
}
export const LIST_PLAN_LIMITS = Object.freeze({
  maxDepth: 12,
  maxNodes: 100_000,
  maxItems: 10_000,
  maxBytes: 2 * 1024 * 1024,
});
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const named = (node: OoxmlNode | undefined, name: string) =>
  node && node.kind !== 'textValue'
    ? node.children.find(
        (child) =>
          child.kind !== 'textValue' && child.namespaceUri === W && child.localName === name
      )
    : undefined;

/** Follow native inline containers in Proposed view; historical property trees never enter it. */
function ownedMarkers(paragraph: OoxmlNode): readonly OoxmlNode[] | null {
  if (paragraph.kind === 'textValue') return [];
  const stack = paragraph.children.map((node) => ({ node, depth: 0 })),
    found: OoxmlNode[] = [];
  let count = 0;
  while (stack.length) {
    const { node, depth } = stack.pop()!;
    if (++count > LIST_PLAN_LIMITS.maxNodes || depth > 64) return null;
    if (
      node.kind === 'textValue' ||
      node.kind === 'revisionDelete' ||
      node.kind === 'revisionMoveFrom'
    )
      continue;
    if (
      node.kind === 'contentControl' &&
      contentControlPropertiesOf(node).tag === SCIOBOT_CHECKLIST_TAG
    ) {
      found.push(node);
      continue;
    }
    if (
      node.kind === 'contentControl' ||
      node.kind === 'contentControlContent' ||
      node.kind === 'hyperlink' ||
      node.kind === 'revisionInsert' ||
      node.kind === 'revisionMoveTo' ||
      (node.kind === 'generic' &&
        node.namespaceUri === W &&
        ['customXml', 'smartTag', 'dir', 'bdo'].includes(node.localName))
    )
      for (const child of node.children) stack.push({ node: child, depth: depth + 1 });
  }
  return found;
}

function listContext(pkg: OoxmlPackage) {
  const styleRoot = stylesPartOf(pkg)?.root ?? null,
    relationship = pkg.relationships
      .get(pkg.mainDocumentPart)
      ?.find(
        (record) =>
          record.type ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
      ),
    resolved = relationship && resolveRelationship(relationship),
    part =
      (resolved?.mode === 'Internal' && resolved.target.ok
        ? pkg.parts.get(resolved.target.partName)
        : undefined) ?? pkg.parts.get('/word/numbering.xml'),
    styles = buildStyleCascadeTable(styleRoot);
  return {
    styles,
    part,
    numbering: withNumberingStyleLinks(buildNumberingIndex(part?.root), styles),
  };
}

/** Read a native level's start before assigning a new nested group to an existing instance. */
export function listInstanceStart(pkg: OoxmlPackage, numId: string, level: number): number | null {
  const resolved = listInstanceLevel(pkg, numId, level);
  return resolved ? (resolved.startOverride ?? resolved.level.start ?? 1) : null;
}

/** Resolve imported numbering through its relationship and linked styles. */
export function listInstanceLevel(pkg: OoxmlPackage, numId: string, level: number) {
  return resolveNumberingLevel(listContext(pkg).numbering, numId, level);
}

/** Cache canonical list reads and native style resolution for one immutable package. */
export function createListSourceReader(pkg: OoxmlPackage, part: OoxmlPart) {
  const context = listContext(pkg);
  const indexed = (name: string, key: string) => {
    const values = new Map<string, OoxmlNode>();
    const duplicates = new Set<string>();
    for (const node of context.part?.root.children ?? []) {
      if (node.kind === 'textValue' || node.namespaceUri !== W || node.localName !== name) continue;
      const value = node.attributes.find((a) => a.namespaceUri === W && a.localName === key)?.value;
      if (!value || duplicates.has(value)) continue;
      if (values.has(value)) {
        values.delete(value);
        duplicates.add(value);
      } else values.set(value, node);
    }
    return values;
  };
  let numbers: ReturnType<typeof indexed> | undefined,
    abstracts: ReturnType<typeof indexed> | undefined;
  const levelShapes = new Map<string, string | null>();
  const attr = (node: OoxmlNode | undefined, name: string) =>
    node && node.kind !== 'textValue'
      ? node.attributes.find((a) => a.namespaceUri === W && a.localName === name)?.value
      : undefined;
  const levelIn = (node: OoxmlNode | undefined, name: string, level: number) =>
    node && node.kind !== 'textValue'
      ? node.children.find(
          (child) =>
            child.kind !== 'textValue' &&
            child.namespaceUri === W &&
            child.localName === name &&
            attr(child, 'ilvl') === String(level)
        )
      : undefined;
  const cache = new Map<string, ReturnType<typeof inspectListSource>>();
  const paragraphs = new Map<string, ReturnType<typeof resolveParagraphLayoutInputs>>();
  return {
    numbering: context.numbering,
    read(paragraphId: string) {
      let value = cache.get(paragraphId);
      if (!value) {
        value = inspectListSource(part, paragraphId, context.styles);
        cache.set(paragraphId, value);
      }
      return value;
    },
    levelShape(numId: string, level: number) {
      const key = `${numId}:${level}`;
      if (levelShapes.has(key)) return levelShapes.get(key)!;
      numbers ??= indexed('num', 'numId');
      abstracts ??= indexed('abstractNum', 'abstractNumId');
      const num = numbers.get(numId);
      const abstract = abstracts.get(attr(named(num, 'abstractNumId'), 'val') ?? '');
      const node =
        named(levelIn(num, 'lvlOverride', level), 'lvl') ?? levelIn(abstract, 'lvl', level);
      let shape: string | null = null;
      if (node && num && num.kind !== 'textValue' && abstract && abstract.kind !== 'textValue') {
        const template = {
          ...abstract,
          kind: 'generic' as const,
          attributes: abstract.attributes.filter(
            (a) => !(a.namespaceUri === W && a.localName === 'abstractNumId')
          ),
        };
        const concrete = {
          ...num,
          kind: 'generic' as const,
          attributes: num.attributes.filter(
            (a) => !(a.namespaceUri === W && a.localName === 'numId')
          ),
          children: num.children.flatMap((child): OoxmlNode[] => {
            if (child.kind === 'textValue' || child.namespaceUri !== W) return [child];
            if (child.localName === 'abstractNumId') return [];
            if (child.localName !== 'lvlOverride') return [child];
            const children = child.children.filter(
              (field) =>
                !(
                  field.kind !== 'textValue' &&
                  field.namespaceUri === W &&
                  field.localName === 'startOverride' &&
                  field.children.length === 0 &&
                  field.attributes.every((a) => a.namespaceUri === W && a.localName === 'val')
                )
            );
            return children.length ||
              child.attributes.some((a) => a.namespaceUri !== W || a.localName !== 'ilvl')
              ? [{ ...child, children }]
              : [];
          }),
        };
        shape = JSON.stringify([
          canonicalOoxmlFingerprint(template),
          canonicalOoxmlFingerprint(concrete),
        ]);
      }
      levelShapes.set(key, shape);
      return shape;
    },
    level(numId: string, level: number) {
      return resolveNumberingLevel(context.numbering, numId, level);
    },
    paragraph(paragraphId: string) {
      let value = paragraphs.get(paragraphId);
      if (!value) {
        const node = findNode(part, paragraphId);
        if (node?.kind !== 'paragraph') return null;
        value = resolveParagraphLayoutInputs(node, 1, context.styles);
        paragraphs.set(paragraphId, value);
      }
      return value;
    },
  };
}

function inspectListSource(
  part: OoxmlPart,
  paragraphId: string,
  styles: ReturnType<typeof listContext>['styles']
): NativeEditFailure | { readonly success: true; readonly state: ListSourceState } {
  const paragraph = findNode(part, paragraphId);
  if (paragraph?.kind !== 'paragraph')
    return fail('The native list paragraph does not exist.', 'invalid-document');
  const pPr = propertyContainer(paragraph, 'paragraphProperties', 'pPr');
  const native = readNumPr(cascadeParagraphFormatting(styles, pPr).paragraphPropertyNodes);
  const membership =
    native && native.numId !== '0' ? { numId: native.numId, level: native.ilvl } : null;
  const owned = ownedMarkers(paragraph);
  if (!owned) return fail('The checklist marker scan exceeds its limit.', 'invalid-document');
  if (owned.length > 1)
    return fail('The paragraph has multiple owned checklist markers.', 'ambiguous-target');
  const marker = owned[0],
    markerProperties = marker && contentControlPropertiesOf(marker);
  if (markerProperties && (!markerProperties.checkbox || markerProperties.dataBinding))
    return fail('The owned checklist marker is invalid or data-bound.', 'unsupported-operation');
  const indentation = named(pPr, 'ind');
  let checkboxLevel = 0;
  if (marker && indentation?.kind !== 'textValue') {
    const value = (name: string) =>
      indentation?.attributes.find(
        (attribute) => attribute.namespaceUri === W && attribute.localName === name
      )?.value;
    const left = value('start') ?? value('left');
    const twips = left === undefined ? 720 : readTwipsMeasure(left);
    if (
      twips === null ||
      twips === undefined ||
      !Number.isInteger(twips / 720) ||
      twips < 720 ||
      twips > 6480
    )
      return fail(
        'The owned checklist indentation is not a supported list level.',
        'unsupported-operation'
      );
    checkboxLevel = twips / 720 - 1;
  }
  const state: ListSourceState = {
    numbering: membership,
    checkbox: marker
      ? {
          controlId: marker.id,
          checked: markerProperties!.checkbox!.checked,
          level: checkboxLevel,
        }
      : null,
    fingerprint: sha256FontBytes(
      new TextEncoder().encode(
        JSON.stringify([
          named(pPr, 'numPr') ? canonicalOoxmlFingerprint(named(pPr, 'numPr')!) : null,
          indentation ? canonicalOoxmlFingerprint(indentation) : null,
          marker ? canonicalOoxmlFingerprint(marker) : null,
        ])
      )
    ),
  };
  return { success: true, state };
}

/**
 * Plan membership only. Displayed ordinals and marker strings never enter the input or output.
 * New instance keys join items in this plan; executors allocate IDs inside their transaction.
 * Use the same native snapshot for every patch. Text and sibling insertion use D02/D05 separately.
 */
export function planListMembership(
  source: ListPlanSource,
  patches: readonly ListMembershipPatch[]
): EditPlanResult {
  const problem = jsonDataProblem(patches, LIST_PLAN_LIMITS);
  if (problem) return fail(problem);
  if (!Array.isArray(patches) || patches.length > LIST_PLAN_LIMITS.maxItems)
    return fail('Invalid or excessive list patches.');
  const raw = JSON.stringify(patches);
  if (
    raw.length > LIST_PLAN_LIMITS.maxBytes ||
    new TextEncoder().encode(raw).byteLength > LIST_PLAN_LIMITS.maxBytes
  )
    return fail('List patches exceed their byte limit.');
  const token = createEditPlan(source?.revision, []);
  if (!token.success) return token;
  const part = source.pkg?.parts.get(source.partName);
  if (!part) return fail('The native list part does not exist.', 'invalid-document');
  const reader = createListSourceReader(source.pkg, part);
  const numbering = reader.numbering;
  const seen = new Set<string>(),
    newInstances = new Map<string, string>(),
    newStyles = new Map<string, string | null>(),
    operations: CanonicalEdit[] = [];
  for (const patch of patches) {
    if (
      !patch ||
      typeof patch !== 'object' ||
      Array.isArray(patch) ||
      typeof patch.paragraphId !== 'string' ||
      Object.keys(patch).some((key) => key !== 'paragraphId' && key !== 'target') ||
      !isListTarget(patch.target)
    )
      return fail('Invalid list membership patch.');
    if (seen.has(patch.paragraphId))
      return fail('A paragraph has duplicate list patches.', 'ambiguous-target');
    seen.add(patch.paragraphId);
    const inspected = reader.read(patch.paragraphId);
    if (!inspected.success) return inspected;
    const state = inspected.state;
    const membership = state.numbering,
      marker = state.checkbox;
    const target = patch.target;
    if (target.kind === 'none' && !membership && !marker) continue;
    if (target.kind === 'checklist') {
      if (
        !membership &&
        state.checkbox &&
        state.checkbox.checked === target.checked &&
        state.checkbox.level === target.level
      )
        continue;
    } else if (target.kind !== 'none') {
      if (target.instance.kind === 'existing') {
        const resolved = resolveNumberingLevel(numbering, target.instance.numId, target.level);
        const format = resolved?.level.numFmt;
        if (
          !format ||
          format === 'none' ||
          (target.kind === 'bullet' ? format !== 'bullet' : format === 'bullet')
        )
          return fail(
            'The target list instance has no matching format at this level.',
            'unsupported-operation'
          );
        if (
          !marker &&
          membership &&
          membership.numId === target.instance.numId &&
          membership.level === target.level
        )
          continue;
      } else {
        const style = target.style ?? null;
        if (newStyles.has(target.instance.key) && newStyles.get(target.instance.key) !== style)
          return fail('One new list instance has conflicting marker styles.', 'ambiguous-target');
        newStyles.set(target.instance.key, style);
        const levelKey = JSON.stringify([target.instance.key, target.level]),
          definition = JSON.stringify([
            target.kind,
            target.style ?? null,
            target.kind === 'ordered' ? (target.instance.start ?? 1) : null,
          ]);
        const previous = newInstances.get(levelKey);
        if (previous !== undefined && previous !== definition)
          return fail('One new list key has conflicting definitions.', 'ambiguous-target');
        newInstances.set(levelKey, definition);
      }
    }
    operations.push({
      kind: 'list',
      partName: part.name,
      paragraphId: patch.paragraphId,
      before: state,
      after: target,
    });
  }
  return createEditPlan(source.revision, operations);
}
