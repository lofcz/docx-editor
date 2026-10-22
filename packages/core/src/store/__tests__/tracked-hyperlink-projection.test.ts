import { expect, test } from 'bun:test';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

test('current-owned hyperlink projection preserves 200000 siblings on V8', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'docx-link-projection-'));
  try {
    const outfile = join(directory, 'projection.mjs');
    const build = await Bun.build({
      entrypoints: [
        fileURLToPath(new URL('../store/tree-op-tracked-link-structure.ts', import.meta.url)),
      ],
      target: 'node',
      format: 'esm',
    });
    expect(build.success).toBe(true);
    await Bun.write(outfile, build.outputs[0]!);
    // Bun optimizes push spreads. V8 exposes the argument-count limit in supported Node hosts.
    const script = `
      import { proposedHyperlinkContent } from ${JSON.stringify(pathToFileURL(outfile).href)};
      import assert from 'node:assert/strict';
      const nodes = Array.from({length:200000}, (_,i) => ({kind:'run', id:'r'+i, children:[]}));
      const tail = {kind:'run', id:'tail', children:[]};
      const deleted = {kind:'revisionDelete', children:[{kind:'run', id:'gone', children:[]}]};
      const actual = proposedHyperlinkContent([
        {kind:'revisionInsert', children:nodes}, deleted,
        {kind:'revisionInsert', children:[deleted, {kind:'revisionInsert', children:[tail]}]}
      ]);
      assert.equal(actual.length, nodes.length+1);
      for(let i=0;i<nodes.length;i++) assert.equal(actual[i],nodes[i]);
      assert.equal(actual.at(-1),tail);
    `;
    const { NODE_OPTIONS: _inheritedNodeOptions, ...env } = process.env;
    const child = spawnSync('node', ['--input-type=module', '--eval', script], {
      env,
      encoding: 'utf8',
      timeout: 30000,
    });
    expect(child.status, child.stderr).toBe(0);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
