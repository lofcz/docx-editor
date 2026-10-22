// Development-only oracle. Execute the application's actual accepted/rejected property views.
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const app = resolve(process.argv[2] ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next');
const cases = [
  {
    name: 'rename',
    before: { name: 'Old', duration: 10 },
    after: { name: 'New', duration: 10 },
    keys: ['name'],
  },
  {
    name: 'duration and resource',
    before: { name: 'Old', duration: 10, resource: 'A' },
    after: { name: 'Old', duration: 15, resource: 'B' },
    keys: ['duration', 'resource'],
  },
  {
    name: 'add absent name',
    before: { duration: 10 },
    after: { name: 'New', duration: 10 },
    keys: ['name'],
  },
  {
    name: 'remove duration',
    before: { name: 'Old', duration: 10 },
    after: { name: 'Old' },
    keys: ['duration'],
  },
  {
    name: 'null to zero',
    before: { name: 'Old', duration: null },
    after: { name: 'Old', duration: 0 },
    keys: ['duration'],
  },
  {
    name: 'zero to null',
    before: { name: 'Old', duration: 0 },
    after: { name: 'Old', duration: null },
    keys: ['duration'],
  },
  {
    name: 'remove explicit null',
    before: { name: null, duration: 10 },
    after: { duration: 10 },
    keys: ['name'],
  },
  {
    name: 'add explicit null',
    before: { duration: 10 },
    after: { name: null, duration: 10 },
    keys: ['name'],
  },
];
const program = `
import { getCommittedBlocks, getWorkingBlocks } from './src/lib/suggestion-views.ts';
const cases = JSON.parse(${JSON.stringify(JSON.stringify(cases))});
const output = cases.map(value => {
  const pick = source => Object.fromEntries(value.keys.filter(key => Object.hasOwn(source,key)).map(key => [key,source[key]]));
  const node = { type: 'lesson_activity', ...value.after, children: [{type:'p',children:[{text:'Body'}]}],
    suggestion: { type:'update', properties:pick(value.before), newProperties:pick(value.after) } };
  const attrs = node => Object.fromEntries(Object.entries(node).filter(([key]) => !['type','children'].includes(key)));
  return { ...value, committed:attrs(getCommittedBlocks([node])[0]), working:attrs(getWorkingBlocks([node])[0]) };
});
process.stdout.write(JSON.stringify(output));
`;
const output = JSON.parse(execFileSync('bun', ['--eval', program], { cwd: app, encoding: 'utf8' }));
const source = 'src/lib/suggestion-views.ts';
const fixture = {
  applicationRevision: execFileSync('git', ['rev-parse', 'HEAD'], {
    cwd: app,
    encoding: 'utf8',
  }).trim(),
  sources: {
    [source]: createHash('sha256')
      .update(readFileSync(resolve(app, source)))
      .digest('hex'),
  },
  cases: output,
};
writeFileSync(
  resolve(root, 'packages/sciobot/src/__tests__/fixtures/domain-properties.json'),
  `${JSON.stringify(fixture, null, 2)}\n`
);
process.stdout.write(`${output.length} application property-view cases captured.\n`);
