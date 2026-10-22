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
import { CUSTOM_ALLOWED_TAGS, CUSTOM_LITERAL_TAGS, CUSTOM_DATA_ONLY_TAGS } from './src/components/chat/mdx-tags.ts';
import { sharedMdxMarkdown } from './src/components/mdx/rules.ts';
import { unified } from 'unified';
import remarkParse from 'remark-parse';
import remarkMdx from 'remark-mdx';
import remarkGfm from 'remark-gfm';
import remarkMath from 'remark-math';
const parser = unified().use(remarkParse).use(remarkMath).use(remarkGfm).use(remarkMdx);
function firstJsx(node) {
  if (node.type === 'mdxJsxFlowElement' || node.type === 'mdxJsxTextElement') return node;
  for (const child of node.children ?? []) {
    const found = firstJsx(child);
    if (found) return found;
  }
}
function capture(id, markdown) {
  const ast = JSON.parse(JSON.stringify(firstJsx(parser.parse(markdown))));
  const type = ast.type === 'mdxJsxTextElement' ? 'shared_mdx_inline' : 'shared_mdx_block';
  const element = { type, mdast: ast, children: [{ text: '' }] };
  return { id, markdown, type, ast, serialized: sharedMdxMarkdown(element).trimEnd() };
}
const cases = [];
for (const [tag, names] of Object.entries(CUSTOM_ALLOWED_TAGS)) {
  const attributes = names.map((name) => name + '="' + tag + '.' + name + ' &amp; &quot;quoted&quot; 😀"').join(' ');
  const open = '<' + tag + (attributes ? ' ' + attributes : '') + '>';
  const close = '</' + tag + '>';
  const component = open + 'Stored 😀 ' + tag + close;
  cases.push(capture(tag + '-block', open + '\\n\\nStored 😀 ' + tag + '\\n\\n' + close));
  cases.push(capture(tag + '-inline', 'Before ' + component + ' after.'));
}
cases.push(capture('nested-listening', '<listening title="Two speakers" language="cs" seed="42">\\n\\n<speaker id="anna" name="Anna" voice="warm" />\\n\\n<line speaker="anna">Listen **carefully** 😀.</line>\\n\\n<speaker id="jan" name="Jan" portrait-id="portrait-1" />\\n\\n<line speaker="jan">{JSON.stringify({answer: [1, 2]})}</line>\\n\\n</listening>'));
cases.push(capture('nested-suggestions', '<suggestions>\\n\\n<suggestion message="Continue">Next <competence code="ABC" framework="RVP">Goal</competence>.</suggestion>\\n\\n<suggestion message="Stop">{({safe: true})}</suggestion>\\n\\n</suggestions>'));
process.stdout.write(JSON.stringify({ allowedTags: CUSTOM_ALLOWED_TAGS, literalTags: CUSTOM_LITERAL_TAGS, dataOnlyTags: CUSTOM_DATA_ONLY_TAGS, cases }));
`,
    ],
    { cwd: app, encoding: 'utf8', maxBuffer: 1024 * 1024 }
  )
);
const files = [
  'src/components/chat/mdx-tags.ts',
  'src/components/mdx/rules.ts',
  'supabase/functions/_shared/activities/audio-exercise/markup.ts',
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
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/mdx-tags.json'),
  JSON.stringify(fixture, null, 2) + '\n'
);
writeFileSync(
  resolve(root, 'packages/sciobot/src/projection/mdx-tags.ts'),
  `// Captured from application CUSTOM_ALLOWED_TAGS. Regenerate with capture-mdx-reference.mjs.\nexport const CUSTOM_ALLOWED_TAGS: Readonly<Record<string, readonly string[]>> = ${JSON.stringify(data.allowedTags, null, 2)};\nexport const CUSTOM_LITERAL_TAGS: readonly string[] = ${JSON.stringify(data.literalTags)};\nexport const CUSTOM_DATA_ONLY_TAGS: readonly string[] = ${JSON.stringify(data.dataOnlyTags)};\n`
);
process.stdout.write(
  `Captured ${Object.keys(data.allowedTags).length} custom MDX tags from application exports.\n`
);
