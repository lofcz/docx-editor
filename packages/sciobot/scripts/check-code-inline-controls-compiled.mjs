// Run the original acceptance fixtures against the built private SDK and one built core.
// Support fixtures remain test-only source. The built evaluation transport uses the EigenPal Pro License.
import { existsSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const packageRoot = fileURLToPath(new URL('../', import.meta.url));
const repoRoot = resolve(packageRoot, '../..');
const packages = join(repoRoot, 'packages');
const paths = {
  '@docx-editor.dev/core': [join(packages, 'core/dist/index.js')],
  '@docx-editor.dev/core/*': [join(packages, 'core/dist/*.js')],
  '@docx-editor.dev/pro/collaboration': [join(packages, 'pro/dist/collaboration/index.js')],
  '@docx-editor.dev/docx-to-markdown': [join(packages, 'docx-to-markdown/dist/index.js')],
  '@docx-editor.dev/docx-to-markdown/*': [join(packages, 'docx-to-markdown/dist/*.js')],
  '@docx-editor.dev/fonts': [join(packages, 'fonts/dist/index.js')],
  '@sciobot/docx': [join(packageRoot, 'dist/index.js')],
  '@sciobot/docx/collaboration': [join(packageRoot, 'dist/collaboration/index.js')],
};
for (const [name, [entry]] of Object.entries(paths)) {
  if (!name.endsWith('/*') && !existsSync(entry)) {
    throw Error(
      `Build core, fonts, docx-to-markdown, and sciobot before running: missing ${entry}`
    );
  }
}
const directory = mkdtempSync(join(tmpdir(), 'sciobot-code-controls-compiled-'));
try {
  writeFileSync(
    join(directory, 'tsconfig.json'),
    JSON.stringify({
      extends: join(repoRoot, 'tsconfig.json'),
      compilerOptions: { paths },
    })
  );
  writeFileSync(
    join(directory, 'bunfig.toml'),
    `[test]\npreload = [${JSON.stringify(join(repoRoot, 'scripts/test/preload-dom.ts'))}]\n`
  );
  symlinkSync(join(packageRoot, 'node_modules'), join(directory, 'node_modules'), 'dir');
  const collaboration = join(directory, 'collaboration.ts');
  const support = readFileSync(join(packageRoot, 'test/support/collaboration.ts'), 'utf8')
    .replace("from '../../src/collaboration'", "from '@sciobot/docx/collaboration'")
    .replace(
      "from '../../../pro/src/collaboration/document-session'",
      "from '@docx-editor.dev/pro/collaboration'"
    )
    .replace(
      "from '../../src/native/prepare-review-groups'",
      `from '${join(packageRoot, 'src/native/prepare-review-groups')}'`
    );
  if (/from ['"]\.{1,2}\//.test(support)) throw Error('Unmapped collaboration fixture import.');
  writeFileSync(collaboration, support);

  const selected = process.argv.slice(2);
  if (selected.some((name) => !/^(?:review\/|core\/)?[a-z0-9-]+\.test\.ts$/.test(name)))
    throw Error('Pass diff fixture basenames, review/name.test.ts, or core/name.test.ts.');
  const fixtures = selected.length
    ? selected
    : [
        'code-inline-atoms.test.ts',
        'code-inline-atoms-collaboration.test.ts',
        'code-atom-middle-peer-undo.test.ts',
      ];
  for (const filename of fixtures) {
    let source = readFileSync(
      filename.startsWith('core/')
        ? join(packages, 'core/src/store/__tests__', filename.slice(5))
        : join(packageRoot, filename.startsWith('review/') ? 'src' : 'src/diff', filename),
      'utf8'
    );
    if (filename.startsWith('core/'))
      source = source.replace(
        /from '(?:\.\.\/package\/[^']+|\.\.\/store\/(?:tree-ops|tree-op-types)\.ts)'/g,
        "from '@docx-editor.dev/core/store'"
      );
    source = source.replace(
      /from '(?:\.\.\/(?:review|projection|blocks|ai|native)\/[^']+|\.\.\/diff\/(?:plan-target|plan|target-key)|\.\/(?:plan-target|plan|target-key))'/g,
      "from '@sciobot/docx'"
    );
    if (filename.startsWith('review/'))
      source = source.replace(
        /from '\.\/(?:native-module|recompose|working-view|committed-view|groups|accept-all|reject-all)'/g,
        "from '@sciobot/docx'"
      );
    source = source.replaceAll(
      "from '../../test/support/collaboration'",
      `from '${collaboration}'`
    );
    source = source.replaceAll(
      "from '../../test/support/",
      `from '${join(packageRoot, 'test/support')}/`
    );
    // Runtime product imports must use the compiled package. Keep every test body unchanged.
    if (/from ['"]\.{1,2}\//.test(source)) throw Error(`Unmapped product import in ${filename}`);
    writeFileSync(join(directory, filename.replace('/', '--')), source);
  }
  const result = spawnSync(
    'bun',
    [
      'test',
      '--tsconfig-override',
      join(directory, 'tsconfig.json'),
      ...fixtures.map((filename) => join(directory, filename.replace('/', '--'))),
    ],
    { cwd: directory, stdio: 'inherit' }
  );
  if (result.error) throw result.error;
  process.exitCode = result.status ?? 1;
} finally {
  rmSync(directory, { recursive: true, force: true });
}
