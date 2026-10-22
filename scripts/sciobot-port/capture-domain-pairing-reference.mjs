// Development-only oracle. Run the application's actual strategies and installed Plate diff engine.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const child = (text) => [{ type: 'p', children: [{ text }] }];
const pair = (name, type, before = {}, after = {}) => ({
  name,
  before: { type, ...before, children: child('Old question') },
  after: { type, ...before, ...after, children: child('New question') },
});
const cases = [
  pair('phase child edit', 'lesson_phase', { phaseType: 'evocation' }),
  pair(
    'phase type replacement',
    'lesson_phase',
    { phaseType: 'evocation' },
    { phaseType: 'reflection' }
  ),
  pair(
    'recap phase replacement',
    'recap_phase',
    { phaseType: 'evocation' },
    { phaseType: 'reflection' }
  ),
  pair('title grade replacement', 'lesson_title', { grade: '7' }, { grade: '8' }),
  pair(
    'deep equal identity',
    'lesson_title',
    { grade: { label: '7', track: ['A'] } },
    { grade: { track: ['A'], label: '7' } }
  ),
  pair(
    'curriculum source replacement',
    'recap_curriculum',
    { framework: 'RVP', source: 'catalog' },
    { source: 'school' }
  ),
  ...['preparation_recap', 'recap_goals', 'recap_materials', 'recap_adaptations'].map((type) =>
    pair(type + ' children', type)
  ),
  pair('recap own property replacement', 'recap_goals', { extra: 'old' }, { extra: 'new' }),
  pair(
    'activity name and duration',
    'lesson_activity',
    { name: 'Warmup', duration: 10 },
    { name: 'Opening', duration: 15 }
  ),
  pair(
    'activity resource only',
    'lesson_activity',
    { name: 'Warmup', duration: 10, resource: 'A' },
    { resource: 'B' }
  ),
  pair(
    'activity resource with duration',
    'lesson_activity',
    { name: 'Warmup', duration: 10, resource: 'A' },
    { resource: 'B', duration: 15 }
  ),
  pair(
    'ignored generated properties',
    'lesson_activity',
    { id: 'old', listStart: 2, name: 'Warmup' },
    { id: 'new', listStart: 3 }
  ),
  pair('source excerpt children', 'source_excerpt', {
    url: 'https://a.example',
    trust: 'verified',
  }),
  pair(
    'source excerpt citation replacement',
    'source_excerpt',
    { url: 'https://a.example', trust: 'verified' },
    { url: 'https://b.example' }
  ),
  ...['lesson_visual', 'listening', 'source'].map((type) => pair(type + ' atomic', type)),
];
const program = `
import { computeDiff } from '@platejs/diff';
import { lessonPlanDiffStrategy } from './src/components/lesson-plan/diff-strategy.ts';
import { sourceDiffStrategy } from './src/components/source/diff-strategy.ts';
const cases = JSON.parse(${JSON.stringify(JSON.stringify(cases))});
const strategy = node => lessonPlanDiffStrategy(node) ?? sourceDiffStrategy(node);
const output = cases.map(value => {
  const result = computeDiff([value.before], [value.after], {
    granularity: 'block', runScopeWordHints: true, ignoreProps: ['id', 'listStart'], isInline: () => false,
    getDeleteProps: () => ({ change: 'delete' }), getInsertProps: () => ({ change: 'insert' }),
    getUpdateProps: (_node, old, next) => ({ update: { old, next } }), getDiffStrategy: strategy,
  });
  return { ...value, strategy: strategy(value.before), retainsContainer: strategy(value.before)?.kind === 'container' && result.length === 1 && !result[0].change,
    rootChanges: result.flatMap(node => node.change ? [node.change] : []),
    update: result.length === 1 ? result[0].update ?? null : null };
});
process.stdout.write(JSON.stringify(output));
`;
const output = JSON.parse(
  execFileSync('bun', ['--eval', program], {
    cwd: app,
    encoding: 'utf8',
    maxBuffer: 8 * 1024 * 1024,
  })
);
const files = [
  'src/components/lesson-plan/diff-strategy.ts',
  'src/components/source/diff-strategy.ts',
  'node_modules/@platejs/diff/dist/index.js',
];
const fixture = {
  schemaVersion: 1,
  source: {
    revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: app, encoding: 'utf8' }).trim(),
    files: Object.fromEntries(
      files.map((file) => [
        file,
        createHash('sha256')
          .update(readFileSync(resolve(app, file)))
          .digest('hex'),
      ])
    ),
  },
  cases: output,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/domain-pairing.json'),
  JSON.stringify(fixture, null, 2) + '\n'
);
process.stdout.write(`Captured ${output.length} application container decisions.\n`);
