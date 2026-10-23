// Original shared-history assertions against compiled core and document API packages.
// The EigenPal Pro License evaluation transport remains test-only source.
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const repoRoot = fileURLToPath(new URL('../../../', import.meta.url));
const packages = join(repoRoot, 'packages');
const paths = {
  '@docx-editor.dev/core': [join(packages, 'core/dist/index.js')],
  '@docx-editor.dev/core/*': [join(packages, 'core/dist/*.js')],
  '@docx-editor.dev/editor-api': [join(packages, 'editor-api/dist/index.mjs')],
  '@docx-editor.dev/editor-api/*': [join(packages, 'editor-api/dist/*.mjs')],
};
for (const [name, [entry]] of Object.entries(paths))
  if (!name.endsWith('/*') && !existsSync(entry))
    throw Error(`Build core and editor-api before running: missing ${entry}`);
const directory = mkdtempSync(join(tmpdir(), 'sciobot-joined-paragraphs-compiled-'));
try {
  writeFileSync(
    join(directory, 'tsconfig.json'),
    JSON.stringify({ extends: join(repoRoot, 'tsconfig.json'), compilerOptions: { paths } })
  );
  writeFileSync(
    join(directory, 'bunfig.toml'),
    `[test]\npreload = [${JSON.stringify(join(repoRoot, 'scripts/test/preload-dom.ts'))}]\n`
  );
  symlinkSync(join(packages, 'sciobot/node_modules'), join(directory, 'node_modules'), 'dir');
  const sourceRoot = join(packages, 'pro/src');
  const filename = 'document-joined-paragraph-automation.test.ts';
  const source = readFileSync(join(sourceRoot, 'collaboration/__tests__', filename), 'utf8')
    .replace("from '../../index.ts'", `from '${join(sourceRoot, 'index.ts')}'`)
    .replace(
      "from './document-peer-support.ts'",
      `from '${join(sourceRoot, 'collaboration/__tests__/document-peer-support.ts')}'`
    );
  if (/from ['"]\.{1,2}\//.test(source)) throw Error('Unmapped relative fixture import');
  writeFileSync(join(directory, filename), source);
  const result = spawnSync(
    'bun',
    ['test', '--tsconfig-override', join(directory, 'tsconfig.json'), join(directory, filename)],
    { cwd: directory, stdio: 'inherit' }
  );
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(directory, { recursive: true, force: true });
}
