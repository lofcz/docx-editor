// Development only. Execute the application's resolver against its installed Plate fork.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(
  process.env.SCIOBOT_REFERENCE_APP ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next'
);
const quotePath = resolve(app, 'src/lib/quote-focus.ts');
const quoteSource = readFileSync(quotePath, 'utf8');
const file = ts.createSourceFile(quotePath, quoteSource, ts.ScriptTarget.Latest, true);
const names = new Set([
  'isRemoveSuggestion',
  'translateLivePathToWorking',
  'resolveQuoteSelection',
]);
const declarations = file.statements.filter(
  (node) =>
    (ts.isFunctionDeclaration(node) && names.has(node.name?.text)) ||
    (ts.isVariableStatement(node) &&
      node.declarationList.declarations.some((entry) => names.has(entry.name.getText(file))))
);
if (declarations.length !== names.size)
  throw new Error('The source resolver declarations changed.');
const testSource = readFileSync(resolve(app, 'src/lib/ai-review-diff.test.ts'), 'utf8');
const editorFile = ts.createSourceFile('reference.ts', testSource, ts.ScriptTarget.Latest, true);
let markdownConfig;
const findConfig = (node) => {
  if (
    ts.isCallExpression(node) &&
    node.expression.getText(editorFile) === 'MarkdownPlugin.configure'
  )
    markdownConfig = node.getText(editorFile);
  ts.forEachChild(node, findConfig);
};
findConfig(editorFile);
if (!markdownConfig) throw new Error('The C01 editor Markdown configuration changed.');
const editorSource = testSource
  .slice(0, testSource.indexOf('describe('))
  .replace(markdownConfig, '...MarkdownKit')
  .replace(/import \{ describe, expect, it \} from "@rstest\/core";\s*/, '')
  .replace(
    /from ["']([^"']+)["']/g,
    (_, specifier) =>
      `from ${JSON.stringify(
        specifier.startsWith('@/')
          ? resolve(app, 'src', specifier.slice(2))
          : specifier.startsWith('.')
            ? resolve(app, 'src/lib', specifier)
            : Bun.resolveSync(specifier, resolve(app, 'src/lib'))
      )}`
  );
const casePath = resolve(root, 'scripts/sciobot-port/quote-reference-cases.json');
const inputs = JSON.parse(readFileSync(casePath, 'utf8'));
const temporary = mkdtempSync(resolve(tmpdir(), 'sciobot-quote-reference-'));
try {
  const resultPath = resolve(temporary, 'result.json');
  const runner = resolve(temporary, 'capture.ts');
  writeFileSync(
    runner,
    `
import { writeFileSync } from 'node:fs';
import { resolveSelectionByPath } from ${JSON.stringify(Bun.resolveSync('@platejs/markdown', resolve(app, 'src/lib')))};
import { getWorkingBlocks } from ${JSON.stringify(resolve(app, 'src/lib/suggestion-views.ts'))};
import { MarkdownKit } from ${JSON.stringify(resolve(app, 'src/components/markdown-kit.tsx'))};
${editorSource}
${declarations.map((node) => node.getText(file)).join('\n')}
const cases = [];
for (const input of ${JSON.stringify(inputs)}) {
  const editor = editorFor(input.review?.before ?? '');
  if (input.review) editor.tf.setValue(prepareAIReviewDiff(editor, input.review.after));
  else editor.tf.setValue(input.value);
  let selection = input.selection;
  if (input.selectText) {
    const find = (nodes, prefix = []) => {
      for (const [index, node] of nodes.entries()) {
        const path = [...prefix, index];
        if (typeof node.text === 'string' && node.text.includes(input.selectText)) return { path, offset: node.text.indexOf(input.selectText) };
        const child = Array.isArray(node.children) ? find(node.children, path) : null;
        if (child) return child;
      }
      return null;
    };
    const anchor = find(editor.children);
    if (!anchor) throw new Error('Missing selected source text: ' + input.name);
    selection = { anchor, focus: { path: anchor.path, offset: anchor.offset + input.selectText.length } };
  }
  const expected = resolveQuoteSelection(editor, selection);
  const working = getWorkingBlocks(editor.children);
  const textOf = (nodes) => nodes.map((node) => typeof node.text === 'string' ? node.text : textOf(node.children ?? [])).join('');
  cases.push({ ...input, selection, workingText: textOf(working), expected });
}
writeFileSync(${JSON.stringify(resultPath)}, JSON.stringify(cases));
`
  );
  execFileSync('bun', ['--tsconfig-override', resolve(app, 'tsconfig.json'), runner], {
    cwd: app,
    stdio: 'inherit',
  });
  const files = [
    'src/lib/quote-focus.ts',
    'src/lib/suggestion-views.ts',
    'src/lib/ai-review-diff.test.ts',
    'src/components/lesson-plan/md-rules.ts',
    'src/components/source/md-rules.ts',
    'src/components/markdown-kit.tsx',
    'src/components/mdx/rules.ts',
    'node_modules/@platejs/markdown/dist/index.js',
  ];
  const artifact = {
    schemaVersion: 1,
    source: {
      revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: app, encoding: 'utf8' }).trim(),
      files: Object.fromEntries(
        files.map((path) => [
          path,
          createHash('sha256')
            .update(readFileSync(resolve(app, path)))
            .digest('hex'),
        ])
      ),
      inputHash: createHash('sha256').update(readFileSync(casePath)).digest('hex'),
      scope:
        'The actual source resolver and working view run with the C01 reference editor plugins and production MarkdownKit. UI and stores are not loaded.',
    },
    cases: JSON.parse(readFileSync(resultPath, 'utf8')),
  };
  writeFileSync(
    resolve(root, 'packages/sciobot/src/__tests__/fixtures/quote-reference.json'),
    `${JSON.stringify(artifact, null, 2)}\n`
  );
  process.stdout.write(`Captured ${artifact.cases.length} source quote results.\n`);
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
