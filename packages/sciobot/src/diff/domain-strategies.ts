/** Application rules ported from lesson-plan and source diff-strategy.ts. */
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

const STRATEGIES: Readonly<Record<string, DomainDiffStrategy>> = {
  lesson_title: { kind: 'container', identityProps: ['grade'] },
  preparation_recap: { kind: 'container' },
  recap_goals: { kind: 'container' },
  recap_curriculum: { kind: 'container', identityProps: ['framework', 'source'] },
  recap_materials: { kind: 'container' },
  recap_adaptations: { kind: 'container' },
  recap_phase: { kind: 'container', identityProps: ['phaseType'] },
  lesson_phase: { kind: 'container', identityProps: ['phaseType'] },
  lesson_activity: { kind: 'container', updateProps: ['name', 'duration'] },
  lesson_visual: { kind: 'atomic' },
  listening: { kind: 'atomic' },
  callout: { kind: 'container' },
  story: { kind: 'container' },
  solution: { kind: 'container' },
  source: { kind: 'atomic' },
  source_excerpt: {
    kind: 'container',
    identityProps: ['url', 'title', 'author', 'bookId', 'trust'],
  },
};

export const IGNORED_DIFF_PROPERTIES: ReadonlySet<string> = new Set(['id', 'listStart']);

export function domainDiffStrategy(node: DomainNode): DomainDiffStrategy | undefined {
  return node.type && Object.hasOwn(STRATEGIES, node.type) ? STRATEGIES[node.type] : undefined;
}

export function sameContainerIdentity(before: DomainNode, after: DomainNode): boolean {
  if (before.type !== after.type) return false;
  const strategy = domainDiffStrategy(before);
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
  after: DomainNode
): readonly DomainPropertyChange[] {
  const strategy = domainDiffStrategy(before);
  if (strategy?.kind !== 'container' || !sameContainerIdentity(before, after)) return [];
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
