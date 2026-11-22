import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { statSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

// Run against the installed tarball. The bundler must follow its public export map.
const cwd = path.resolve(process.argv[2]);
const entry = path.join(cwd, 'grapheme-probe.mjs');
const output = path.join(cwd, 'grapheme-bundle.mjs');
writeFileSync(entry, "export { segmentGraphemes } from '@docx-editor.dev/core/layout';\n");
execFileSync('bun', ['build', entry, '--target', 'browser', '--minify', '--outfile', output], {
  cwd,
  stdio: 'inherit',
});
assert.ok(statSync(output).size < 4_096, 'A single grapheme reader pulled in unrelated code');
const { segmentGraphemes } = await import(pathToFileURL(output));
assert.deepEqual(
  segmentGraphemes('a👩🏾‍💻').map((segment) => segment.text),
  ['a', '👩🏾‍💻']
);

// The layout algorithms and the standalone reader must share their segmentation state.
writeFileSync(entry, "export * from '@docx-editor.dev/core/layout';\n");
const layout = await import(pathToFileURL(entry));
let calls = 0;
layout.setGraphemeBoundary({
  segment(text) {
    calls++;
    return [{ index: 0, utf16From: 0, utf16To: text.length, text }];
  },
});
try {
  assert.equal(layout.segmentGraphemes('a b').length, 1);
  assert.equal(layout.boundedFallbackWordSegments('c d').length, 1);
  assert.equal(calls, 2);
} finally {
  layout.resetGraphemeBoundary();
}
