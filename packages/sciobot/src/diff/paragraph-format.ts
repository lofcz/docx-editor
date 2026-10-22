import { directParagraphProperties, findNode, type OoxmlPart } from '@docx-editor.dev/core/store';
import type { NativeEditFailure, RevisionToken } from '../contracts/document';
import { jsonDataProblem, type JsonValue } from '../data/json';
import {
  createEditPlan,
  type EditPlanResult,
  type PlannedPropertyChange,
  type PlannedValue,
} from './plan';

export interface ParagraphFormatSource {
  readonly revision: RevisionToken;
  readonly part: OoxmlPart;
  readonly paragraphId: string;
}
export interface ParagraphPropertyPatch {
  readonly localName: string;
  /** Merge named native attributes. Null removes an attribute; omitted attributes stay unchanged. */
  readonly attributes?: Readonly<Record<string, string | null>>;
}
export interface ParagraphFormatTarget {
  readonly properties: readonly ParagraphPropertyPatch[];
  /** Remove complete properties to restore inheritance. Numbering and tabs use separate planners. */
  readonly remove?: readonly string[];
}
const properties = new Set([
  'pStyle',
  'jc',
  'spacing',
  'ind',
  'keepNext',
  'keepLines',
  'widowControl',
  'pageBreakBefore',
  'contextualSpacing',
  'bidi',
  'shd',
]);
const invalidXml = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;
const loneSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const state = (value: JsonValue | undefined): PlannedValue =>
  value === undefined ? { present: false } : { present: true, value };
const fail = (
  error: string,
  code: NativeEditFailure['code'] = 'invalid-input'
): NativeEditFailure => ({ success: false, code, error });

/**
 * Plan only named paragraph properties over authored native values. The source stays unchanged.
 * Values use native XML units and style IDs. Callers resolve styles before planning.
 * Attribute patches preserve independent settings within spacing, indentation, and shading.
 * Null can remove a conflicting hanging/firstLine attribute without clearing other indentation.
 * Executors validate revisions and old values, then merge deltas over the complete live native bag.
 */
export function planParagraphFormat(
  source: ParagraphFormatSource,
  target: ParagraphFormatTarget
): EditPlanResult {
  const problem = jsonDataProblem(target, { maxDepth: 8, maxNodes: 4096 });
  if (problem) return fail(problem);
  if (
    !target ||
    !Array.isArray(target.properties) ||
    (target.remove !== undefined && !Array.isArray(target.remove)) ||
    Object.keys(target).some((key) => key !== 'properties' && key !== 'remove')
  )
    return fail('Invalid paragraph property target.');
  const names = new Set<string>();
  for (const property of target.properties) {
    if (
      !property ||
      typeof property !== 'object' ||
      Array.isArray(property) ||
      !properties.has(property.localName) ||
      names.has(property.localName) ||
      Object.keys(property).some((key) => key !== 'localName' && key !== 'attributes')
    )
      return fail('Unsupported or duplicate paragraph property.');
    names.add(property.localName);
    if (
      property.attributes !== undefined &&
      (!property.attributes ||
        typeof property.attributes !== 'object' ||
        Array.isArray(property.attributes) ||
        Object.entries(property.attributes).some(
          ([key, value]) =>
            !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(key) ||
            (value !== null &&
              (typeof value !== 'string' || invalidXml.test(value) || loneSurrogate.test(value)))
        ))
    )
      return fail('Invalid paragraph property attributes.');
  }
  for (const name of target.remove ?? []) {
    if (typeof name !== 'string' || !properties.has(name) || names.has(name))
      return fail('Unsupported or duplicate removed paragraph property.');
    names.add(name);
  }
  if (
    !source ||
    typeof source.paragraphId !== 'string' ||
    findNode(source.part, source.paragraphId)?.kind !== 'paragraph'
  )
    return fail('The native paragraph does not exist.', 'invalid-document');
  const authored = new Map(
    directParagraphProperties(source.part, source.paragraphId).map((property) => [
      property.localName,
      { ...property.attributes },
    ])
  );
  const changes: PlannedPropertyChange[] = [];
  for (const patch of target.properties as readonly ParagraphPropertyPatch[]) {
    const before = authored.get(patch.localName);
    const after: Record<string, string> = { ...before };
    for (const [key, value] of Object.entries(patch.attributes ?? {})) {
      if (value === null) delete after[key];
      else after[key] = value;
    }
    changes.push({ name: patch.localName, before: state(before), after: state(after) });
  }
  for (const name of target.remove ?? [])
    changes.push({ name, before: state(authored.get(name)), after: { present: false } });
  return createEditPlan(source.revision, [
    {
      kind: 'format',
      partName: source.part.name,
      target: { kind: 'paragraph', paragraphId: source.paragraphId },
      changes,
    },
  ]);
}
