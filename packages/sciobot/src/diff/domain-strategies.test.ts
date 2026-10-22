import { expect, test } from 'bun:test';
import {
  domainDiffStrategy,
  domainPropertyChanges,
  sameContainerIdentity,
} from './domain-strategies';

test('activity name and duration updates keep the container paired', () => {
  const before = { type: 'lesson_activity', name: 'Warmup', duration: 10, children: ['Original'] };
  const after = { ...before, name: 'Opening', duration: 15, children: ['Edited'] };
  expect(sameContainerIdentity(before, after)).toBe(true);
  expect(domainPropertyChanges(before, after)).toEqual([
    {
      key: 'name',
      before: { present: true, value: 'Warmup' },
      after: { present: true, value: 'Opening' },
    },
    { key: 'duration', before: { present: true, value: 10 }, after: { present: true, value: 15 } },
  ]);
});

test('identity changes replace phase, title, curriculum, and citation wrappers', () => {
  const fixtures = [
    [
      { type: 'lesson_phase', phaseType: 'evocation' },
      { type: 'lesson_phase', phaseType: 'reflection' },
    ],
    [
      { type: 'recap_phase', phaseType: 'evocation' },
      { type: 'recap_phase', phaseType: 'reflection' },
    ],
    [
      { type: 'lesson_title', grade: '7' },
      { type: 'lesson_title', grade: '8' },
    ],
    [
      { type: 'recap_curriculum', framework: 'RVP', source: 'catalog' },
      { type: 'recap_curriculum', framework: 'RVP', source: 'school' },
    ],
    [
      { type: 'source_excerpt', url: 'https://a.example' },
      { type: 'source_excerpt', url: 'https://b.example' },
    ],
    [
      { type: 'source_excerpt', trust: 'verified' },
      { type: 'source_excerpt', trust: 'generated' },
    ],
  ];
  for (const [before, after] of fixtures)
    expect(sameContainerIdentity(before!, after!)).toBe(false);
});

test('absent and explicitly undefined properties remain distinct', () => {
  const before = { type: 'lesson_activity' };
  expect(domainPropertyChanges(before, { ...before, duration: 5 })).toEqual([
    { key: 'duration', before: { present: false }, after: { present: true, value: 5 } },
  ]);
  expect(domainPropertyChanges({ ...before, duration: undefined }, before)).toEqual([
    { key: 'duration', before: { present: true, value: undefined }, after: { present: false } },
  ]);
});

test('generated IDs and list renumbering do not create property changes', () => {
  const before = { type: 'lesson_activity', id: 'old', listStart: 2, name: 'Warmup', duration: 10 };
  const after = { ...before, id: 'new', listStart: 3 };
  expect(sameContainerIdentity(before, after)).toBe(true);
  expect(domainPropertyChanges(before, after)).toEqual([]);
});

test('all attribute-free recap wrappers retain identity through child edits', () => {
  for (const type of ['preparation_recap', 'recap_goals', 'recap_materials', 'recap_adaptations']) {
    expect(sameContainerIdentity({ type, children: ['Old'] }, { type, children: ['New'] })).toBe(
      true
    );
  }
});

test('visuals, listening exercises, and source chips stay atomic', () => {
  for (const type of ['lesson_visual', 'listening', 'source']) {
    expect(domainDiffStrategy({ type })).toEqual({ kind: 'atomic' });
    expect(sameContainerIdentity({ type }, { type })).toBe(false);
  }
  expect(domainDiffStrategy({ type: 'paragraph' })).toBeUndefined();
  expect(domainDiffStrategy({ type: 'constructor' })).toBeUndefined();
});
