import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const read = (file) => readFileSync(resolve(app, file), 'utf8').replaceAll('\r\n', '\n');
const bridge = read('src/hooks/chat/useWorkspaceEditBridge.ts');
const server = read('supabase/functions/ai-chat_shared/tools/document-edit-tools.ts');
function between(source, start, end) {
  const first = source.indexOf(start);
  const last = source.indexOf(end, first + start.length);
  if (first === -1 || last === -1) throw new Error(`Reference boundary changed: ${start}`);
  return source.slice(first, last);
}
const aiPrefix = read('src/lib/ai-review-diff.test.ts')
  .split('describe(')[0]
  .replace(/import \{ describe, expect, it \} from "@rstest\/core";\s*/, '')
  .replace(
    /from ["']([^"']+)["']/g,
    (_, name) =>
      `from ${JSON.stringify(
        name.startsWith('@/')
          ? resolve(app, 'src', name.slice(2))
          : Bun.resolveSync(name, resolve(app, 'src/lib'))
      )}`
  );
const clientReadLimit = Number(bridge.match(/const READ_DOCUMENT_DEFAULT_LIMIT = (\d+)/)?.[1]);
const declaredReadLimit = Number(server.match(/const DEFAULT_READ_LIMIT = (\d+)/)?.[1]);
const maxSearchResults = Number(server.match(/const MAX_SEARCH_RESULTS = (\d+)/)?.[1]);
if (![clientReadLimit, declaredReadLimit, maxSearchResults].every(Number.isFinite))
  throw new Error('Missing limits');
const helpers = between(bridge, 'function readMarkdown(', '/**\n * Bridges');
const runEdit = between(bridge, 'function runEdit(', 'interface SetInput');
const working = between(
  bridge,
  'const serializeWorkingFromEditor =',
  'const ensureLessonSession ='
);
const executor = between(bridge, 'const buildPlateExecutor =', '// ----- Worksheet executor');
const factory = between(
  server,
  'export function createDocumentTools(',
  'export { DEFAULT_READ_LIMIT'
);
const imports = [
  ['{ serializeMd }', '@platejs/markdown'],
  ['{ z }', './supabase/functions/_shared/schema.ts'],
  ['{ findAndReplaceMarkdown }', './src/lib/agent-edit-find-replace.ts'],
  ['{ analyzeEditTagBalance, buildEditResultEcho }', './src/lib/agent-edit-feedback.ts'],
  [
    '{ formatEditRejection, formatSchemaWarning, validateLessonPlanMarkdown, validateLessonPlanSyntax }',
    './src/lib/lesson-plan-validation.ts',
  ],
  ['{ normalizeMdxLineBreaks }', './supabase/functions/_shared/normalize-mdx-line-breaks.ts'],
  ['{ getWorkingBlocks, getCommittedBlocks }', './src/lib/suggestion-views.ts'],
]
  .map(([names, name]) => `import ${names} from ${JSON.stringify(Bun.resolveSync(name, app))};`)
  .join('\n');

const temp = mkdtempSync(resolve(tmpdir(), 'sciobot-tool-oracle-'));
try {
  const resultFile = resolve(temp, 'result.json');
  const runner = resolve(temp, 'capture.ts');
  writeFileSync(
    runner,
    `
import { writeFileSync } from 'node:fs';
${aiPrefix}
${imports}
const READ_DOCUMENT_DEFAULT_LIMIT = ${clientReadLimit};
const DEFAULT_READ_LIMIT = ${declaredReadLimit};
const MAX_SEARCH_RESULTS = ${maxSearchResults};
const tool = (options) => options;
${helpers}
${runEdit}
${factory}
const paths = ['lesson', 'document'];
const tools = createDocumentTools(paths);
const schemaInputs = [
  ['read_document', { path: 'lesson' }],
  ['read_document', { path: 'missing' }],
  ['read_document', { path: 'lesson', offset: 0 }],
  ['search_document', { path: 'document', pattern: 'Alpha' }],
  ['search_document', { path: 'lesson', pattern: 'Alpha', max_results: 31 }],
  ['edit_document', { path: 'document', old_string: 'Alpha', new_string: 'Beta' }],
  ['edit_document', { path: 'document', old_string: 'Alpha', new_string: 'Beta', is_suggesion: true }],
  ['edit_document', { path: 'document', old_string: 'Alpha' }],
  ['edit_document', { path: 'document', old_string: 'Alpha', new_string: 'Beta', expected_replacements: 0 }],
  ['set_document', { path: 'document', markdown: '# New' }],
  ['set_document', { path: 'lesson', markdown: '# New', is_suggesion: true }],
  ['set_document', { path: 'missing', markdown: '# New' }],
];
const schemaCases = schemaInputs.map(([toolName, input]) => {
  const parsed = tools[toolName].inputSchema.safeParse(input);
  return { tool: toolName, input, expected: parsed.success
    ? { success: true, data: parsed.data }
    : { success: false, issues: parsed.error.issues.map(({ code, path, message }) => ({ code, path, message })) } };
});
const readCases = [
  { name: 'default', projection: 'Alpha\\nBeta\\nGamma', input: {} },
  { name: 'window', projection: 'Alpha\\nBeta\\nGamma', input: { offset: 2, limit: 1 } },
  { name: 'past end', projection: 'Alpha\\nBeta\\nGamma', input: { offset: 9 } },
  { name: 'effective client default', projection: Array.from({ length: 1002 }, (_, i) => 'Line ' + (i + 1)).join('\\n'), input: {} },
].map((item) => ({ ...item, expected: readMarkdown(item.projection, item.input) }));
const searchCases = [
  { name: 'literal', projection: 'Alpha[a]\\nBeta\\nalpha[a]', input: { pattern: 'Alpha[a]', isRegex: false, maxResults: 20 } },
  { name: 'regex limit', projection: 'Alpha\\nBeta\\nALPHA', input: { pattern: '^alpha', isRegex: true, maxResults: 1 } },
  { name: 'invalid regex', projection: 'Alpha', input: { pattern: '[', isRegex: true, maxResults: 20 } },
  { name: 'empty pattern', projection: 'Alpha', input: { pattern: '', isRegex: false, maxResults: 20 } },
].map((item) => ({ ...item, expected: searchMarkdown(item.projection, item.input) }));
function capture(item) {
  const editor = editorFor(item.beforeMarkdown);
  const streamWorkingMdRef = { current: null };
  const workingSnapshotRef = { current: null };
  const firstApplyDiffOfStreamRef = { current: true };
  const effects = [];
  const ensureLessonSession = () => {};
  const review = {
    applyDirect(markdown) {
      effects.push('direct');
      const { before, value } = prepareAIDirectEdit(editor, markdown);
      applyAnimatedAIEdit(editor, value, { beforeValue: before, reducedMotion: true });
    },
    applyDiff(markdown) {
      effects.push('suggestion');
      editor.tf.setValue(prepareAIReviewDiff(editor, markdown));
    },
  };
  ${working}
  ${executor}
  const one = buildPlateExecutor(item.materialKind === 'lesson_plan');
  const wire = item.input;
  const input = item.tool === 'edit_document' ? {
    oldString: wire.old_string, newString: wire.new_string,
    replaceAll: wire.replace_all === true, expectedReplacements: wire.expected_replacements,
    isSuggestion: wire.is_suggesion === true,
  } : { markdown: wire.markdown, isSuggestion: wire.is_suggesion === true };
  const result = item.tool === 'edit_document' ? one.edit(input) : one.set(input);
  return { ...item, expected: { result, effects,
    workingMarkdown: serializeMd(editor, { value: getWorkingBlocks(editor.children) }),
    committedMarkdown: serializeMd(editor, { value: getCommittedBlocks(editor.children) }) } };
}
const executorCases = [
  { name: 'direct default', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha.\\n\\nTail.', input: { path: 'document', old_string: 'Alpha.', new_string: '**Beta.**' } },
  { name: 'explicit direct', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', old_string: 'Alpha.', new_string: 'Beta.', is_suggesion: false } },
  { name: 'explicit suggestion', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', old_string: 'Alpha.', new_string: 'Beta.', is_suggesion: true } },
  { name: 'ambiguous edit', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha. Alpha.', input: { path: 'document', old_string: 'Alpha.', new_string: 'Beta.' } },
  { name: 'count mismatch', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha. Alpha.', input: { path: 'document', old_string: 'Alpha.', new_string: 'Beta.', replace_all: true, expected_replacements: 1 } },
  { name: 'malformed edit', tool: 'edit_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', old_string: 'Alpha.', new_string: '<activity>Broken' } },
  { name: 'incremental shape warning', tool: 'edit_document', materialKind: 'lesson_plan', beforeMarkdown: 'Alpha.', input: { path: 'lesson', old_string: 'Alpha.', new_string: 'Beta.' } },
  { name: 'set generic', tool: 'set_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', markdown: '# Replacement' } },
  { name: 'set suggestion', tool: 'set_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', markdown: '# Replacement', is_suggesion: true } },
  { name: 'set invalid syntax', tool: 'set_document', materialKind: 'document', beforeMarkdown: 'Alpha.', input: { path: 'document', markdown: '<activity>Broken' } },
  { name: 'set invalid lesson schema', tool: 'set_document', materialKind: 'lesson_plan', beforeMarkdown: 'Alpha.', input: { path: 'lesson', markdown: '# Replacement' } },
].map(capture);
let emptyPathsError;
try { createDocumentTools([]); } catch (error) { emptyPathsError = error.message; }
writeFileSync(${JSON.stringify(resultFile)}, JSON.stringify({
  paths, defaults: { declaredReadLimit: DEFAULT_READ_LIMIT, clientReadLimit: READ_DOCUMENT_DEFAULT_LIMIT,
    maxSearchResults: MAX_SEARCH_RESULTS }, emptyPathsError, schemaCases, readCases, searchCases, executorCases,
}));
`
  );
  execFileSync('bun', ['--tsconfig-override', resolve(app, 'tsconfig.json'), runner], {
    cwd: app,
    stdio: 'inherit',
  });
  const cases = JSON.parse(readFileSync(resultFile, 'utf8'));
  const files = Object.fromEntries(
    [
      'src/hooks/chat/useWorkspaceEditBridge.ts',
      'src/lib/agent-edit-feedback.ts',
      'src/lib/lesson-plan-validation.ts',
      'supabase/functions/ai-chat_shared/tools/document-edit-tools.ts',
    ].map((file) => [
      file,
      createHash('sha256')
        .update(readFileSync(resolve(app, file)))
        .digest('hex'),
    ])
  );
  const revision = execFileSync('git', ['rev-parse', 'HEAD'], {
    cwd: app,
    encoding: 'utf8',
  }).trim();
  const destination = resolve(root, 'packages/sciobot/src/__tests__/fixtures/tool-contracts.json');
  mkdirSync(dirname(destination), { recursive: true });
  writeFileSync(
    destination,
    `${JSON.stringify(
      {
        schemaVersion: 1,
        source: {
          revision,
          files,
          scope:
            'Source-executed schemas, read/search functions, and executor closure. Identity SDK tool wrapper; immediate reference review delegate. Path enum is exercised; mounted hook routing belongs to I06.',
        },
        ...cases,
      },
      null,
      2
    )}\n`
  );
  process.stdout.write(
    `Captured ${cases.schemaCases.length} schema and ${cases.executorCases.length} executor examples.\n`
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
