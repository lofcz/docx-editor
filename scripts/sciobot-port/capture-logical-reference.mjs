// Development only. Execute the existing application, without importing replacement code.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const source = readFileSync(resolve(app, 'src/lib/ai-review-diff.test.ts'), 'utf8');
const prefix = source
  .slice(0, source.indexOf('describe('))
  .replace(/import \{ describe, expect, it \} from "@rstest\/core";\s*/, '')
  .replace(
    /from ["']([^"']+)["']/g,
    (_, specifier) =>
      `from ${JSON.stringify(specifier.startsWith('@/') ? resolve(app, 'src', specifier.slice(2)) : Bun.resolveSync(specifier, resolve(app, 'src/lib')))}`
  );
const cases = [
  ['inline title', '<lesson_title grade="7">Lesson</lesson_title>'],
  ['flow title', '<lesson_title grade="7">\n\nLesson\n\n</lesson_title>'],
  ['standalone source', '<source title="Book" url="https://example.org" type="verified" />'],
  [
    'source excerpt',
    '<source title="Book" url="https://example.org">\n\nQuoted words.\n\n</source>',
  ],
  ['empty paired source', '<source title="Book" url="https://example.org"></source>'],
  ['recap goals lists', '<recap_goals>\n\n3. First goal\n4. Second goal\n\n</recap_goals>'],
  [
    'recap goals legacy prefixes',
    '<recap_goals>\n\n1\\. First goal\n\n2\\. Second goal\n\n</recap_goals>',
  ],
  ['recap curriculum', '<recap_curriculum framework="RVP">\n\n- One\n- Two\n\n</recap_curriculum>'],
  ['recap materials', '<recap_materials>\n\n- One\n- Two\n\n</recap_materials>'],
  ['recap adaptations', '<recap_adaptations>\n\n- One\n- Two\n\n</recap_adaptations>'],
  ['ordinary activity list', '<activity name="A" duration="5">\n\n3. One\n4. Two\n\n</activity>'],
  ['list continuation', '3. First\n\n   Continuation.\n\n   - Child\n\n4. Second'],
];
const temp = mkdtempSync(resolve(tmpdir(), 'sciobot-logical-reference-'));
try {
  const runner = resolve(temp, 'capture.ts'),
    out = resolve(temp, 'result.json');
  writeFileSync(
    runner,
    `import {writeFileSync} from 'node:fs';\n${prefix}\nconst cases=${JSON.stringify(cases)};\nconst output=cases.map(([name,markdown])=>{const editor=editorFor(markdown);const original=withoutIds(editor.children);const normalized=withoutIds(prepareAIDirectEdit(editor,markdown).before);return {name,markdown,original,normalized};});\nwriteFileSync(${JSON.stringify(out)},JSON.stringify(output));\n`
  );
  execFileSync('bun', ['--tsconfig-override', resolve(app, 'tsconfig.json'), runner], {
    cwd: app,
    stdio: 'inherit',
  });
  const files = [
    'src/lib/ai-review-diff.test.ts',
    'src/lib/ai-review-diff.ts',
    'src/components/lesson-plan/md-rules.ts',
    'src/components/source/md-rules.ts',
    'src/lib/lesson-plan-deserialize.ts',
    'node_modules/@platejs/markdown/package.json',
    'node_modules/@platejs/markdown/dist/index.js',
  ];
  const fixture = {
    schemaVersion: 1,
    source: {
      revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: app, encoding: 'utf8' }).trim(),
      files: Object.fromEntries(
        files.map((name) => [
          name,
          createHash('sha256')
            .update(readFileSync(resolve(app, name)))
            .digest('hex'),
        ])
      ),
    },
    cases: JSON.parse(readFileSync(out, 'utf8')),
  };
  writeFileSync(
    resolve(root, 'packages/sciobot/src/__tests__/fixtures/logical-blocks.json'),
    JSON.stringify(fixture, null, 2) + '\n'
  );
  process.stdout.write(
    `Captured ${fixture.cases.length} independently executed application normalization cases.\n`
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
