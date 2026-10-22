import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const data = JSON.parse(
  execFileSync(
    'bun',
    [
      '--eval',
      `
import { calloutRules } from './src/components/callout-md-rules.ts';
import { storyRules } from './src/components/story/md-rules.ts';
import { solutionRules } from './src/components/solution/md-rules.ts';
import { createSlateEditor } from 'platejs';
const editor = createSlateEditor();
const rules = { ...calloutRules, ...storyRules, ...solutionRules };
const cases = [
  { id: 'tip', type: 'callout', attributes: { variant: 'tip', icon: '💡', extra: 'unknown' } },
  { id: 'story', type: 'story', attributes: { zeta: 'first', audience: 'class', count: 2, enabled: true, hidden: false, omitted: null, _custom: 'kept', suggestion: 'inert' } },
  { id: 'solution', type: 'solution', attributes: { answer: '42', details: { nested: ['data', 2] }, constructor: 'inert', onClick: 'globalThis.X37 = true' } },
  { id: 'default', type: 'callout', attributes: { icon: '📝' } },
  { id: 'empty', type: 'callout', attributes: { variant: '', extra: 'kept' } },
  { id: 'numeric', type: 'callout', attributes: { variant: 42 } },
];
for (const item of cases) {
  item.serializedAttributes = rules[item.type].serialize({ type: item.type, ...item.attributes, children: [] }, { editor }).attributes;
}
const aliases = [
  [{ name: 'type', type: 'mdxJsxAttribute', value: ' TIP ' }],
  [{ name: 'variant', type: 'mdxJsxAttribute', value: ' WARNING ' }, { name: 'icon', type: 'mdxJsxAttribute', value: '⚠' }],
  [{ name: 'type', type: 'mdxJsxAttribute', value: 'note' }, { name: 'variant', type: 'mdxJsxAttribute', value: 'tip' }],
].map(attributes => ({ attributes, node: calloutRules.callout.deserialize({ type: 'mdxJsxFlowElement', name: 'callout', attributes, children: [] }, {}, { editor }) }));
process.stdout.write(JSON.stringify({ cases, aliases }));
`,
    ],
    { cwd: app, encoding: 'utf8' }
  )
);
const files = [
  'src/components/callout-md-rules.ts',
  'src/components/story/md-rules.ts',
  'src/components/solution/md-rules.ts',
  'node_modules/@platejs/markdown/dist/index.js',
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
  ...data,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/container-tags.json'),
  JSON.stringify(fixture, null, 2) + '\n'
);
console.log(
  `Captured ${data.cases.length} application container attribute cases and ${data.aliases.length} aliases.`
);
