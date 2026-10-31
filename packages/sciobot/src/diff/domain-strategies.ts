import type { MdxComponentRegistry } from '../registry/contract';
import { sciobotMdxRegistry } from '../registry/sciobot';

/** How the planner pairs and updates a registered component's nodes. */
export type DomainDiffStrategy =
  | { readonly kind: 'atomic' }
  | {
      readonly kind: 'container';
      readonly identityProps?: readonly string[];
      readonly updateProps?: readonly string[];
    };

export interface DomainNode {
  readonly type?: string;
  readonly [key: string]: unknown;
}

export const IGNORED_DIFF_PROPERTIES: ReadonlySet<string> = new Set(['id', 'listStart']);

/** The registered diff strategy of a node's payload type, if it has one. */
export function domainDiffStrategy(
  node: DomainNode,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): DomainDiffStrategy | undefined {
  return node.type ? registry.diffStrategies.get(node.type) : undefined;
}

export function sameContainerIdentity(
  before: DomainNode,
  after: DomainNode,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): boolean {
  if (before.type !== after.type) return false;
  const strategy = domainDiffStrategy(before, registry);
  if (strategy?.kind !== 'container') return false;
  if (strategy.identityProps)
    return strategy.identityProps.every((key) => equalDomainValue(before[key], after[key]));
  const keys = new Set([...Object.keys(before), ...Object.keys(after)]);
  if (
    [...keys]
      .filter((key) => key !== 'children' && !IGNORED_DIFF_PROPERTIES.has(key))
      .every(
        (key) =>
          Object.hasOwn(before, key) === Object.hasOwn(after, key) &&
          equalDomainValue(before[key], after[key])
      )
  )
    return true;
  // Match the source engine's explicit updatable-property relaxation.
  return (strategy.updateProps ?? []).some(
    (key) =>
      Object.hasOwn(before, key) !== Object.hasOwn(after, key) ||
      !equalDomainValue(before[key], after[key])
  );
}

function equalDomainValue(before: unknown, after: unknown): boolean {
  if (Object.is(before, after)) return true;
  if (
    !before ||
    !after ||
    typeof before !== 'object' ||
    typeof after !== 'object' ||
    Array.isArray(before) !== Array.isArray(after)
  )
    return false;
  const a = before as Record<string, unknown>,
    b = after as Record<string, unknown>;
  const keys = Object.keys(a).filter((key) => !IGNORED_DIFF_PROPERTIES.has(key));
  return (
    keys.length === Object.keys(b).filter((key) => !IGNORED_DIFF_PROPERTIES.has(key)).length &&
    keys.every((key) => Object.hasOwn(b, key) && equalDomainValue(a[key], b[key]))
  );
}

export type PropertyState =
  | { readonly present: false }
  | { readonly present: true; readonly value: unknown };

export interface DomainPropertyChange {
  readonly key: string;
  readonly before: PropertyState;
  readonly after: PropertyState;
}

function propertyState(node: DomainNode, key: string): PropertyState {
  return Object.hasOwn(node, key) ? { present: true, value: node[key] } : { present: false };
}

export function domainPropertyChanges(
  before: DomainNode,
  after: DomainNode,
  registry: MdxComponentRegistry = sciobotMdxRegistry()
): readonly DomainPropertyChange[] {
  const strategy = domainDiffStrategy(before, registry);
  if (strategy?.kind !== 'container' || !sameContainerIdentity(before, after, registry)) return [];
  const changes: DomainPropertyChange[] = [];
  for (const key of strategy.updateProps ?? []) {
    if (IGNORED_DIFF_PROPERTIES.has(key)) continue;
    const oldState = propertyState(before, key);
    const newState = propertyState(after, key);
    if (
      oldState.present === newState.present &&
      (!oldState.present || (newState.present && Object.is(oldState.value, newState.value)))
    )
      continue;
    changes.push({ key, before: oldState, after: newState });
  }
  return changes;
}
