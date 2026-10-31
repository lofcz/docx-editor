import type { DomainDiffStrategy } from '../diff/domain-strategies';

/** Application rules ported from lesson-plan and source diff-strategy.ts, by payload type. */
export const SCIOBOT_DIFF_STRATEGIES: Readonly<Record<string, DomainDiffStrategy>> = {
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
  column_group: { kind: 'container' },
  column: { kind: 'container' },
  story: { kind: 'container' },
  solution: { kind: 'container' },
  toggle: { kind: 'container' },
  date: { kind: 'atomic' },
  toc: { kind: 'atomic' },
  source: { kind: 'atomic' },
  source_excerpt: {
    kind: 'container',
    identityProps: ['url', 'title', 'author', 'bookId', 'trust'],
  },
};
