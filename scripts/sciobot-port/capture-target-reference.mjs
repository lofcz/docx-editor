// Development-only oracle. Execute the production parser plugins from the application checkout.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const pairs = JSON.parse(
  readFileSync(resolve(root, 'packages/sciobot/src/__tests__/fixtures/ai-pairs.json'), 'utf8')
);
const cases = pairs.cases.flatMap((pair) => [
  { name: `${pair.name}:before`, input: pair.beforeMarkdown },
  { name: `${pair.name}:after`, input: pair.afterMarkdown },
]);
cases.push(
  {
    name: 'line breaks preserve code examples',
    input: 'Before<br>after.\n\n`<br>`\n\n```html\n<br>\n```',
  },
  { name: 'interrupting ordered list', input: 'Label:\n6. First\n7. Second' },
  {
    name: 'ordinary numbered prose',
    input: 'The war began\n2. světová válka.\n\nDate:\n15. května.',
  },
  {
    name: 'code and math fences preserve lists',
    input: '```text\nLabel:\n6. First\n7. Second\n```\n\n$$\nLabel:\n6. First\n7. Second\n$$',
  },
  {
    name: 'mentions and encoded identities',
    input:
      '@anna and @jan.\n\n[Display Name](mention:user%20id) and [@kept](https://example.com).\n\n`@code`',
  },
  {
    name: 'gallery runs and linked images',
    input:
      '![One](one.png)\n\n[![Two](two.png)](https://example.com)\n\nText ![Inline](inline.png) stays.\n\n![Three](three.png)',
  },
  {
    name: 'nested image galleries',
    input:
      '<activity name="Gallery" duration="10">\n\n![One](one.png)\n![Two](two.png)\n\nText.\n\n![Three](three.png)\n\n</activity>',
  },
  {
    name: 'opaque nested payloads',
    input:
      '<suggestions>\n\n<suggestion message="Continue">Next <competence code="ABC">Goal</competence>.</suggestion>\n\n</suggestions>\n\n<listening title="Listening"><speaker id="s1" name="Anna" /><line speaker="s1">Hello 😀.</line></listening>',
  },
  {
    name: 'rich Word table cells',
    input:
      '<table><tr><td>\n\n**First**\n\nSecond\n\n<listening title="Audio">\n\n<line speaker="s1">Hello.</line>\n\n</listening>\n\n</td><td>\n\nOther\n\n</td></tr></table>',
  },
  {
    name: 'inert expressions',
    input:
      '<future-widget title={globalThis.__targetProbe()}>\n\n{globalThis.__targetProbe()}\n\n</future-widget>',
  }
);
const failures = ['<activity name="Open">\n\nText.', '<phase>\n\n</activity>', 'Text {broken(}'];
const output = JSON.parse(
  execFileSync(
    'bun',
    [
      '--eval',
      `
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkMath from 'remark-math';
import remarkGfm from 'remark-gfm';
import { remarkMdx, remarkMention } from '@platejs/markdown';
import { remarkImageGalleries } from './src/components/gallery-remark-plugin.ts';
import { normalizeMdxLineBreaks } from './supabase/functions/_shared/normalize-mdx-line-breaks.ts';
import { separateInterruptingLists } from './src/lib/lesson-plan-deserialize.ts';
const processor = unified().use(remarkParse).use(remarkMath).use(remarkGfm).use(remarkMdx).use(remarkMention).use(remarkImageGalleries);
function normalized(value) {
  if (Array.isArray(value)) return value.map(normalized);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key]) => key !== 'position' || typeof value.type !== 'string')
    .map(([key, child]) => [key, normalized(child)])
    .filter(([key, child]) => key !== 'data' || !child || typeof child !== 'object' || Object.keys(child).length));
}
const cases = ${JSON.stringify(cases)}.map(({name, input}) => {
  const markdown = separateInterruptingLists(normalizeMdxLineBreaks(input));
  const tree = processor.runSync(processor.parse(markdown));
  const ast = JSON.parse(JSON.stringify(tree, (key, value) => key === 'estree' ? undefined : value));
  return { name, input, markdown, normalizedAst: normalized(ast) };
});
const failures = ${JSON.stringify(failures)}.map((input) => {
  try { processor.parse(input); throw new Error('Expected the parser to refuse this input'); }
  catch (error) {
    if (!error.reason) throw error;
    return { input, message: error.reason, line: error.line, column: error.column };
  }
});
process.stdout.write(JSON.stringify({ cases, failures }));
`,
    ],
    { cwd: app, encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }
  )
);
const files = [
  'src/components/markdown-kit.tsx',
  'src/components/gallery-remark-plugin.ts',
  'src/lib/lesson-plan-deserialize.ts',
  'supabase/functions/_shared/normalize-mdx-line-breaks.ts',
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
    aiPairsHash: createHash('sha256')
      .update(readFileSync(resolve(root, 'packages/sciobot/src/__tests__/fixtures/ai-pairs.json')))
      .digest('hex'),
  },
  ...output,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/target-parsing.json'),
  JSON.stringify(fixture, null, 2) + '\n'
);
process.stdout.write(
  `Captured ${output.cases.length} target parses and ${output.failures.length} syntax failures.\n`
);
