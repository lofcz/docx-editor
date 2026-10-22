import { segmentGraphemes } from '@docx-editor.dev/core/layout';
import {
  AUTHORABLE_RUN_PROPERTIES,
  authorableHyperlinkTarget,
  authoredProperties,
  findNode,
  hyperlinkTargetOf,
  paragraphTextOf,
  propertyContainer,
  relationshipTargetIn,
  runPropertyEdits,
  runsCovering,
  type FormattingDisplayMode,
  type OoxmlPackage,
  type OoxmlPart,
  type OoxmlProperty,
} from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem, type JsonValue } from '../data/json';
import { createEditPlan, type CanonicalEdit, type EditPlanResult, type PlannedValue } from './plan';

/** Canonical native read inputs. The source stays outside the serializable plan. */
export interface RunFormatSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly paragraphId: string;
  readonly start: number;
  readonly end: number;
  readonly displayMode: FormattingDisplayMode;
}

/** Set only named properties. Removal returns to inheritance; explicit off uses `val: '0'`. */
export interface RunFormatTarget {
  readonly properties: readonly OoxmlProperty[];
  readonly remove?: readonly string[];
}

export interface HyperlinkFormatSource {
  readonly revision: RevisionToken;
  readonly pkg: OoxmlPackage;
  readonly partName: string;
  readonly hyperlinkId: string;
}

/** Omitted fields retain their current values. A null tooltip removes the authored attribute. */
export interface HyperlinkFormatTarget {
  readonly url?: string;
  readonly anchor?: string;
  readonly tooltip?: string | null;
}

const invalidXml = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
const loneSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const xmlText = (value: unknown): value is string =>
  typeof value === 'string' && !invalidXml.test(value) && !loneSurrogate.test(value);
const integer = (value: number) => Number.isSafeInteger(value) && value >= 0;
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });
const state = (value: JsonValue | undefined): PlannedValue =>
  value === undefined ? { present: false } : { present: true, value };

function propertyStates(properties: readonly OoxmlProperty[]): Map<string, JsonValue> {
  return new Map(properties.map((property) => [property.localName, { ...property.attributes }]));
}

function validTarget(target: RunFormatTarget): string | null {
  const problem = jsonDataProblem(target, { maxDepth: 8, maxNodes: 4096 });
  if (problem) return problem;
  if (
    !target ||
    !Array.isArray(target.properties) ||
    (target.remove !== undefined && !Array.isArray(target.remove)) ||
    Object.keys(target).some((key) => key !== 'properties' && key !== 'remove')
  )
    return 'Invalid native run property target.';
  const names = new Set<string>();
  for (const property of target.properties) {
    if (
      !property ||
      typeof property !== 'object' ||
      Array.isArray(property) ||
      !AUTHORABLE_RUN_PROPERTIES.has(property.localName) ||
      names.has(property.localName) ||
      Object.keys(property).some((key) => key !== 'localName' && key !== 'attributes')
    )
      return 'Unsupported or duplicate run property.';
    names.add(property.localName);
    if (
      property.attributes !== undefined &&
      (!property.attributes ||
        typeof property.attributes !== 'object' ||
        Array.isArray(property.attributes) ||
        Object.entries(property.attributes).some(
          ([key, value]) => !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(key) || !xmlText(value)
        ))
    )
      return 'Invalid native run property attributes.';
  }
  for (const name of target.remove ?? []) {
    if (typeof name !== 'string' || !AUTHORABLE_RUN_PROPERTIES.has(name) || names.has(name))
      return 'Unsupported or duplicate removed run property.';
    names.add(name);
  }
  return null;
}

/**
 * Plan property deltas over a canonical range without changing text, relationships, or native state.
 * Native helpers preserve font slots, compound settings, complex-script companions, and review visibility.
 * Formatting field-result atoms needs a separate ownership contract and refuses here.
 * Executors must verify the revision and old values, then merge deltas over live native properties.
 */
export function planRunFormat(source: RunFormatSource, target: RunFormatTarget): EditPlanResult {
  const problem = validTarget(target);
  if (problem) return fail(problem);
  if (
    !source ||
    typeof source.paragraphId !== 'string' ||
    !integer(source.start) ||
    !integer(source.end) ||
    source.start >= source.end ||
    !['all-markup', 'proposed', 'original'].includes(source.displayMode)
  )
    return fail('Invalid native run format range.');
  const paragraph = findNode(source.part, source.paragraphId);
  const text = paragraphTextOf(source.part, source.paragraphId);
  if (paragraph?.kind !== 'paragraph' || text === null)
    return fail('The format paragraph does not exist.', 'invalid-document');
  if (source.end > text.length) return fail('The format range exceeds its paragraph.');
  const boundaries = new Set([
    0,
    text.length,
    ...segmentGraphemes(text).flatMap((item) => [item.utf16From, item.utf16To]),
  ]);
  if (!boundaries.has(source.start) || !boundaries.has(source.end))
    return fail('The format range splits a native grapheme.', 'ambiguous-target');
  const runs = runsCovering(
    source.part,
    source.paragraphId,
    source.start,
    source.end,
    source.displayMode
  );
  const native = runPropertyEdits(
    source.part,
    source.paragraphId,
    source.start,
    source.end,
    target.properties,
    source.displayMode
  );
  if (native.length !== runs.length || native.some((edit) => edit.targetRunIds))
    return fail(
      'Field-result formatting requires explicit run ownership.',
      'unsupported-operation'
    );
  const removed = new Set(target.remove ?? []);
  const edits: CanonicalEdit[] = native.map((edit, index) => {
    const before = propertyStates(
      authoredProperties(
        propertyContainer(runs[index], 'runProperties', 'rPr'),
        AUTHORABLE_RUN_PROPERTIES
      )
    );
    const after = propertyStates(
      edit.properties.filter((property) => !removed.has(property.localName))
    );
    const names = new Set([...before.keys(), ...after.keys()]);
    return {
      kind: 'format',
      partName: source.part.name,
      target: { kind: 'run', paragraphId: source.paragraphId, start: edit.start, end: edit.end },
      changes: [...names].map((name) => ({
        name,
        before: state(before.get(name)),
        after: state(after.get(name)),
      })),
    };
  });
  return createEditPlan(source.revision, edits);
}

/**
 * Plan a target/tooltip change for an existing native hyperlink. No relationship is minted here.
 * Native URL admission remains authoritative. Link creation, removal, and text edits use separate plans.
 */
export function planHyperlinkFormat(
  source: HyperlinkFormatSource,
  target: HyperlinkFormatTarget
): EditPlanResult {
  const problem = jsonDataProblem(target, { maxDepth: 4, maxNodes: 20 });
  if (problem) return fail(problem);
  if (
    !target ||
    Object.keys(target).some((key) => !['url', 'anchor', 'tooltip'].includes(key)) ||
    (target.url !== undefined && target.anchor !== undefined) ||
    (target.url !== undefined && typeof target.url !== 'string') ||
    (target.anchor !== undefined &&
      (!xmlText(target.anchor) || !target.anchor || target.anchor.length > 4096)) ||
    (target.tooltip !== undefined &&
      target.tooltip !== null &&
      (!xmlText(target.tooltip) || target.tooltip.length > 4096))
  )
    return fail('Invalid native hyperlink target.');
  const part = source.pkg.parts.get(source.partName);
  const link = part && findNode(part, source.hyperlinkId);
  if (!part || link?.kind !== 'hyperlink')
    return fail('The native hyperlink does not exist.', 'invalid-document');
  const current = hyperlinkTargetOf(link, (id) =>
    relationshipTargetIn(source.pkg, source.partName, id)
  );
  const changes: Extract<CanonicalEdit, { kind: 'format' }>['changes'][number][] = [];
  if (target.url !== undefined || target.anchor !== undefined) {
    const authored =
      target.url !== undefined ? authorableHyperlinkTarget(target.url) : target.anchor!;
    if (authored === null)
      return fail('The native engine refuses this URL.', 'unsupported-operation');
    const kind = target.url !== undefined ? 'external' : 'internal';
    if (current.kind !== kind || current.authored !== authored)
      changes.push({
        name: 'target',
        before: {
          present: true,
          value: {
            kind: current.kind,
            authored: current.authored,
            ...(current.relationshipId === undefined
              ? {}
              : { relationshipId: current.relationshipId }),
            ...(current.anchor === undefined ? {} : { anchor: current.anchor }),
          },
        },
        after: { present: true, value: { kind, authored } },
      });
  }
  if (Object.hasOwn(target, 'tooltip'))
    changes.push({
      name: 'tooltip',
      before: state(current.tooltip),
      after: state(target.tooltip ?? undefined),
    });
  return createEditPlan(source.revision, [
    {
      kind: 'format',
      partName: source.partName,
      target: { kind: 'hyperlink', hyperlinkId: source.hyperlinkId },
      changes,
    },
  ]);
}
