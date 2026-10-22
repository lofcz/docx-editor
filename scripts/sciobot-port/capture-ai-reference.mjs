// Development only. Execute the application reference, never the replacement.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const plate = resolve(process.argv[3] ?? '/run/media/lofcz/ssd_external/GitHub/plate');
const destination = resolve(root, 'packages/sciobot/src/__tests__/fixtures/ai-pairs.json');
const testFile = resolve(app, 'src/lib/ai-review-diff.test.ts');
const testSource = readFileSync(testFile, 'utf8');
const prefix = testSource.slice(0, testSource.indexOf('describe('));
const arrayStart = testSource.indexOf('of [', testSource.indexOf('describe(')) + 3;
const arrayEnd = testSource.indexOf(']) {', arrayStart) + 1;
if (arrayStart < 3 || arrayEnd < arrayStart) throw new Error('Reference case array changed');
const pairs = testSource.slice(arrayStart, arrayEnd);

function revision(checkout) {
  return execFileSync('git', ['rev-parse', 'HEAD'], { cwd: checkout, encoding: 'utf8' }).trim();
}
function hashes(checkout, files) {
  return Object.fromEntries(
    files.map((file) => [
      file,
      createHash('sha256')
        .update(readFileSync(resolve(checkout, file)))
        .digest('hex'),
    ])
  );
}
const sources = {
  application: {
    revision: revision(app),
    files: hashes(app, [
      'src/lib/ai-review-diff.test.ts',
      'src/lib/ai-review-diff.ts',
      'src/lib/suggestion-views.ts',
      'src/lib/lesson-plan-deserialize.ts',
      'src/components/lesson-plan/md-rules.ts',
      'src/components/lesson-plan/diff-strategy.ts',
      'src/components/source/md-rules.ts',
      'src/components/source/diff-strategy.ts',
      'node_modules/@platejs/ai/package.json',
      'node_modules/@platejs/ai/dist/react/index.js',
      'node_modules/@platejs/suggestion/dist/index.js',
      'node_modules/@platejs/diff/dist/index.js',
    ]),
  },
  plate: {
    revision: revision(plate),
    files: hashes(plate, [
      'packages/ai/src/react/direct-edit/applyAnimatedAIEdit.ts',
      'packages/ai/src/react/direct-edit/planAIEdit.ts',
      'packages/ai/src/react/direct-edit/aiChanges.ts',
      'packages/suggestion/src/lib/diffToSuggestions.ts',
    ]),
  },
};

// Resolve the reference imports from its checkout, including its installed fork.
const imports = prefix
  .replace(/import \{ describe, expect, it \} from "@rstest\/core";\s*/, '')
  .replace(/from ["']([^"']+)["']/g, (_, specifier) => {
    const path = specifier.startsWith('@/')
      ? resolve(app, 'src', specifier.slice(2))
      : Bun.resolveSync(specifier, resolve(app, 'src/lib'));
    return `from ${JSON.stringify(path)}`;
  });
const temp = mkdtempSync(resolve(tmpdir(), 'sciobot-plate-oracle-'));
try {
  const runner = resolve(temp, 'capture.ts');
  writeFileSync(
    runner,
    `
import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
${imports}
const cases = [];
for (const [name, beforeMarkdown, afterMarkdown] of ${pairs}) {
  const direct = editorFor(beforeMarkdown);
  const review = editorFor(beforeMarkdown);
  const original = withoutIds(direct.children);
  review.tf.setValue(prepareAIReviewDiff(review, afterMarkdown));
  acceptAISuggestions(review);
  await Promise.resolve();
  const { before, value } = prepareAIDirectEdit(direct, afterMarkdown);
  const playback = applyAnimatedAIEdit(direct, value, {
    beforeValue: before, reducedMotion: true, session: 'reference-capture',
  });
  await playback.finished;
  const expected = withoutIds(direct.children);
  assert.deepEqual(expected, withoutIds(review.children), name + ': accepted review differs');
  assert.equal(direct.history.undos.length, 1, name + ': direct history is not atomic');
  direct.tf.undo();
  assert.deepEqual(withoutIds(direct.children), original, name + ': undo differs');
  direct.tf.redo();
  assert.deepEqual(withoutIds(direct.children), expected, name + ': redo differs');
  cases.push({ name, beforeMarkdown, afterMarkdown, original,
    canonicalBefore: withoutIds(before), expected });
}
assert.equal(cases.length, 7);
writeFileSync(${JSON.stringify(resolve(temp, 'result.json'))}, JSON.stringify(cases));
`
  );
  execFileSync('bun', ['--tsconfig-override', resolve(app, 'tsconfig.json'), runner], {
    cwd: app,
    stdio: 'inherit',
  });
  const cases = JSON.parse(readFileSync(resolve(temp, 'result.json'), 'utf8'));
  mkdirSync(dirname(destination), { recursive: true });
  writeFileSync(destination, `${JSON.stringify({ schemaVersion: 1, sources, cases }, null, 2)}\n`);
  process.stdout.write(`Captured ${cases.length} independently executed Plate cases.\n`);
} finally {
  rmSync(temp, { recursive: true, force: true });
}
