import type { MdxComponentRegistry } from '../registry/contract';
import { sciobotMdxRegistry } from '../registry/sciobot';
import { diffArrays } from 'diff';
import type { BlockPayload } from '../blocks/payload-schema';
import type { NativeEditFailure } from '../contracts/document';
import type { JsonObject, JsonValue } from '../data/json';
import { MDX_BLOCK, MDX_INLINE, serializeOpaqueMdx } from '../projection/mdx';
import { parseTargetMarkdown, type TargetAstNode } from '../projection/parse-target';
import { planAtomicBlock, type AtomicBlockSource } from './atomic-blocks';
import { domainNodeFromAst } from './domain-ast';
import { payloadMatches } from './inserted-cell-fragments';
import { jsonDataKey, targetAstKey } from './target-key';
import type { EditPlanResult } from './plan';

const fail = (error: string): NativeEditFailure => ({
  success: false,
  code: 'unsupported-operation',
  error,
});
const object = (value: JsonValue | undefined): value is JsonObject =>
  value !== null && typeof value === 'object' && !Array.isArray(value);
const opaque = (payload: BlockPayload) => [MDX_BLOCK, MDX_INLINE].includes(payload.type);

/** Pair atomic syntax without treating its descendants as native editing owners. */
export function atomicTargetsPairable(
  payload: BlockPayload,
  before: TargetAstNode,
  after: TargetAstNode,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  if (opaque(payload))
    return (
      before.type === after.type && ['mdxJsxFlowElement', 'mdxJsxTextElement'].includes(after.type)
    );
  const a = domainNodeFromAst(before, registry),
    b = domainNodeFromAst(after, registry);
  return !!a && !!b && a.type === b.type;
}

function sameShape(a: JsonValue, b: JsonValue): boolean {
  if (!object(a) || !object(b)) return typeof a === typeof b;
  return a.type === b.type && (a.name === undefined || b.name === undefined || a.name === b.name);
}
function syntaxIdentity(value: JsonValue): string | null {
  if (
    object(value) &&
    value.type === 'paragraph' &&
    Array.isArray(value.children) &&
    value.children.length === 1
  ) {
    const nested = syntaxIdentity(value.children[0]!);
    return nested === null ? null : jsonDataKey(['paragraph', nested]);
  }
  if (
    !object(value) ||
    !['mdxJsxFlowElement', 'mdxJsxTextElement'].includes(String(value.type)) ||
    !Array.isArray(value.attributes)
  )
    return null;
  for (const name of ['id', 'key']) {
    const attribute = value.attributes.find(
      (item) => object(item) && item.type === 'mdxJsxAttribute' && item.name === name
    );
    if (object(attribute) && attribute.value !== undefined)
      return jsonDataKey([value.type!, value.name ?? null, name, attribute.value]);
  }
  return null;
}
/** Apply visible AST changes onto stored data, preserving unprojected metadata on retained nodes. */
function mergeAst(stored: JsonValue, before: JsonValue, after: JsonValue): JsonValue {
  if (targetAstKey(before) === targetAstKey(after)) return stored;
  if (Array.isArray(before) && Array.isArray(after)) {
    if (
      !Array.isArray(stored) ||
      stored.length !== before.length ||
      stored.some((node, i) => !sameShape(node, before[i]!))
    )
      throw Error('Opaque syntax needs an unambiguous stored-to-projected child map.');
    const paired = new Map<number, number>();
    const used = new Set<number>();
    const oldIdentities = new Map<string, number[]>(),
      newIdentities = new Map<string, number[]>();
    for (const [values, map] of [
      [before, oldIdentities],
      [after, newIdentities],
    ] as const)
      values.forEach((value, index) => {
        const id = syntaxIdentity(value);
        if (id) map.set(id, [...(map.get(id) ?? []), index]);
      });
    // Explicit unique syntax identities follow a moved child even across exact anchors.
    for (const [id, indices] of newIdentities) {
      const prior = oldIdentities.get(id);
      if (indices.length === 1 && prior?.length === 1) {
        paired.set(indices[0]!, prior[0]!);
        used.add(prior[0]!);
      }
    }

    let oldAt = 0,
      nextAt = 0;
    const removed: number[] = [],
      added: number[] = [];
    const flush = () => {
      // Exact anchors delimit changed sibling windows. Changed matching node kinds retain
      // their private data; removed nodes cannot donate data across an exact anchor.
      let cursor = 0;
      for (const target of added) {
        if (paired.has(target)) continue;
        const match = removed.findIndex((previous, index) => {
          if (index < cursor || used.has(previous) || !sameShape(before[previous]!, after[target]!))
            return false;
          const oldId = syntaxIdentity(before[previous]!),
            nextId = syntaxIdentity(after[target]!);
          return oldId === null || nextId === null || oldId === nextId;
        });
        if (match >= 0) {
          const previous = removed[match]!;
          paired.set(target, previous);
          used.add(previous);
          cursor = match + 1;
        }
      }
      removed.length = 0;
      added.length = 0;
    };
    for (const change of diffArrays(before.map(targetAstKey), after.map(targetAstKey))) {
      if (change.removed) {
        for (let i = 0; i < change.value.length; i++) removed.push(oldAt++);
      } else if (change.added) {
        for (let i = 0; i < change.value.length; i++) added.push(nextAt++);
      } else {
        flush();
        for (let i = 0; i < change.value.length; i++) {
          if (!paired.has(nextAt) && !used.has(oldAt)) {
            paired.set(nextAt, oldAt);
            used.add(oldAt);
          }
          nextAt++;
          oldAt++;
        }
      }
    }
    flush();
    return after.map((value, index) => {
      const previous = paired.get(index);
      return previous === undefined ? value : mergeAst(stored[previous]!, before[previous]!, value);
    });
  }
  if (
    object(stored) &&
    object(before) &&
    object(after) &&
    stored.type === before.type &&
    before.type === after.type
  ) {
    const entries: [string, JsonValue][] = [];
    for (const key of new Set([
      ...Object.keys(stored),
      ...Object.keys(before),
      ...Object.keys(after),
    ])) {
      if (key === 'position' && typeof before.type === 'string') continue;
      if (!Object.hasOwn(before, key)) {
        if (Object.hasOwn(after, key)) entries.push([key, after[key]!]);
        else if (Object.hasOwn(stored, key)) entries.push([key, stored[key]!]);
      } else if (Object.hasOwn(after, key))
        entries.push([
          key,
          Object.hasOwn(stored, key)
            ? mergeAst(stored[key]!, before[key]!, after[key]!)
            : after[key]!,
        ]);
    }
    return Object.fromEntries(entries);
  }
  return after;
}

/** Compile one retained atomic target through D14, retaining identity and opaque storage data. */
export function planTargetAtomic(
  source: AtomicBlockSource,
  payload: BlockPayload,
  before: TargetAstNode,
  after: TargetAstNode,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): EditPlanResult {
  if (!atomicTargetsPairable(payload, before, after, registry))
    return fail('Atomic target syntax changes its native payload kind.');
  try {
    let target: BlockPayload;
    if (opaque(payload)) {
      const projectedBefore = JSON.parse(targetAstKey(before)) as TargetAstNode;
      const projectedAfter = JSON.parse(targetAstKey(after)) as TargetAstNode;
      const mdast = mergeAst(payload.mdast!, projectedBefore, projectedAfter);
      if (!object(mdast)) return fail('An opaque target needs a complete MDX root.');
      target = { ...payload, mdast };
      const parsed = parseTargetMarkdown(serializeOpaqueMdx(target).text);
      if (!parsed.success) return parsed;
      const root = parsed.target.ast.children[0] as TargetAstNode | undefined;
      const syntax =
        payload.type === MDX_INLINE
          ? root && Array.isArray(root.children)
            ? root.children[0]
            : undefined
          : root;
      if (!syntax || targetAstKey(syntax) !== targetAstKey(after))
        return fail('Stored opaque data changes the requested target syntax.');
    } else {
      const definition = registry.blockTags.get(payload.type);
      const old = domainNodeFromAst(before, registry),
        next = domainNodeFromAst(after, registry);
      if (
        definition?.kind !== 'atomic' ||
        !old ||
        !next ||
        !payloadMatches(before, payload, registry)
      )
        return fail('A retained atomic target has no exact registered source codec.');
      if (Array.isArray(after.children) && after.children.length)
        return fail('A registered atomic target needs a codec for its complete child syntax.');
      const attributes = { ...payload.attributes };
      for (const key of new Set([...Object.keys(old), ...Object.keys(next)])) {
        if (
          key === 'type' ||
          (Object.hasOwn(old, key) === Object.hasOwn(next, key) &&
            jsonDataKey(old[key] ?? null) === jsonDataKey(next[key] ?? null))
        )
          continue;
        if (Object.hasOwn(next, key))
          Object.defineProperty(attributes, key, {
            value: next[key],
            enumerable: true,
            writable: true,
            configurable: true,
          });
        else delete attributes[key];
      }
      target = { ...payload, attributes };
      if (!payloadMatches(after, target, registry))
        return fail('The atomic target does not round-trip through its registered codec.');
    }
    return planAtomicBlock(source, target, registry);
  } catch (cause) {
    return fail(cause instanceof Error ? cause.message : String(cause));
  }
}
