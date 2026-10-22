// Development-only oracle. Execute the application's actual echo formatter.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const cases = [
  { name: 'one changed line with context', before: 'a\nb\nc\nd\ne', after: 'a\nb\nCHANGED\nd\ne' },
  { name: 'insert at start', before: 'One\nTwo', after: 'New\nOne\nTwo' },
  { name: 'insert at end', before: 'One\nTwo', after: 'One\nTwo\nNew' },
  { name: 'delete at start', before: 'Deleted\nOne\nTwo', after: 'One\nTwo' },
  { name: 'delete at end', before: 'One\nTwo\nDeleted', after: 'One\nTwo' },
  { name: 'empty landed result', before: 'Deleted', after: '' },
  { name: 'no-op', before: 'One\nTwo\nThree\nFour', after: 'One\nTwo\nThree\nFour' },
  { name: 'unchanged blank lines', before: 'One\n\nOld\n\nTail\n', after: 'One\n\nNew\n\nTail\n' },
  {
    name: 'multiple disjoint replacements',
    before: 'One\nTwo\nThree\nFour\nFive',
    after: 'ONE\nTwo\nThree\nFour\nFIVE',
  },
  {
    name: 'truncated replacement',
    before: 'Old',
    after: Array.from({ length: 100 }, (_, i) => `New ${i + 1}`).join('\n'),
  },
  {
    name: 'exact forty lines',
    before: 'Old',
    after: Array.from({ length: 40 }, (_, i) => `New ${i + 1}`).join('\n'),
  },
  {
    name: 'unicode and escaped markdown',
    before: 'A😀B\n\nOld',
    after: 'A😀B\n\n\\*New\\* &amp; Ž',
  },
  {
    name: 'list ordinals from landed normalization',
    before: '1. First\n\n2. Second',
    after: '1. Updated\n\n2. Second',
  },
];
const expected = JSON.parse(
  execFileSync(
    'bun',
    [
      '--eval',
      `
import { buildEditResultEcho } from './src/lib/agent-edit-feedback.ts';
process.stdout.write(JSON.stringify(${JSON.stringify(cases)}.map((item)=>({...item,message:buildEditResultEcho(item.before,item.after)}))));
`,
    ],
    { cwd: app, encoding: 'utf8' }
  )
);
const files = [
  'src/lib/agent-edit-feedback.ts',
  'src/lib/agent-edit-feedback.test.ts',
  'src/hooks/chat/useWorkspaceEditBridge.ts',
];
const artifact = {
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
  cases: expected,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/edit-echo.json'),
  JSON.stringify(artifact, null, 2) + '\n'
);
process.stdout.write(`Captured ${expected.length} application echo results.\n`);
