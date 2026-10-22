import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const testPath = 'src/lib/agent-edit-find-replace.test.ts';
const source = readFileSync(resolve(app, testPath), 'utf8');
const testNames = [...source.matchAll(/it\("([^"]+)"/g)].map((match) => match[1]);
const destination = resolve(root, 'packages/sciobot/src/__tests__/fixtures/find-replace.json');
const files = Object.fromEntries(
  [
    testPath,
    'src/lib/agent-edit-find-replace.ts',
    'supabase/functions/_shared/find-and-replace.ts',
    'node_modules/diff/package.json',
  ].map((file) => [
    file,
    createHash('sha256')
      .update(readFileSync(resolve(app, file)))
      .digest('hex'),
  ])
);
const temp = mkdtempSync(resolve(tmpdir(), 'sciobot-matcher-oracle-'));
try {
  const runner = resolve(temp, 'capture.ts');
  const body = source
    .replace(/import \{ describe, expect, it \} from "@rstest\/core";\s*/, '')
    .replace(/import \{ findAndReplaceMarkdown \} from "\.\/agent-edit-find-replace";\s*/, '');
  writeFileSync(
    runner,
    `
import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { findAndReplaceMarkdown as referenceFind } from ${JSON.stringify(resolve(app, 'src/lib/agent-edit-find-replace.ts'))};
let testName = '';
const cases = [];
const completed = [];
function describe(_name, fn) { fn(); }
function it(name, fn) { testName = name; fn(); completed.push(name); }
function expect(value) {
  return {
    toBe: (expected) => assert.equal(value, expected),
    toContain: (expected) => assert.ok(value.includes(expected)),
    not: { toContain: (expected) => assert.ok(!value.includes(expected)) },
  };
}
function findAndReplaceMarkdown(content, oldString, newString, replaceAll, expectedReplacements) {
  const expected = referenceFind(content, oldString, newString, replaceAll, expectedReplacements);
  cases.push({ sourceTest: testName,
    input: { content, oldString, newString, replaceAll, expectedReplacements }, expected });
  return expected;
}
${body}
assert.deepEqual(completed, ${JSON.stringify(testNames)});
writeFileSync(${JSON.stringify(resolve(temp, 'result.json'))}, JSON.stringify(cases));
`
  );
  execFileSync('bun', ['--tsconfig-override', resolve(app, 'tsconfig.json'), runner], {
    cwd: app,
    stdio: 'inherit',
  });
  const cases = JSON.parse(readFileSync(resolve(temp, 'result.json'), 'utf8'));
  const revision = execFileSync('git', ['rev-parse', 'HEAD'], {
    cwd: app,
    encoding: 'utf8',
  }).trim();
  mkdirSync(dirname(destination), { recursive: true });
  writeFileSync(
    destination,
    `${JSON.stringify({ schemaVersion: 1, source: { revision, files, testNames }, cases }, null, 2)}\n`
  );
  process.stdout.write(
    `Captured ${cases.length} calls from all ${testNames.length} original matcher tests.\n`
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
