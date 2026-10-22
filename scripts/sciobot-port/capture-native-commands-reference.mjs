// Development-only: execute installed Plate transforms, never native implementation output.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync, rmSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import ts from 'typescript';
import { format, resolveConfig } from 'prettier';
const root = resolve(import.meta.dirname, '../..');
const app = resolve(
  process.env.SCIOBOT_REFERENCE_APP ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next'
);
const temp = mkdtempSync(resolve(tmpdir(), 'native-command-capture-'));
const kits = [
  'indent',
  'basic-blocks',
  'basic-marks',
  'font',
  'list',
  'align',
  'line-height',
  'table',
  'link',
];
const sources = new Set([
  'src/components/workspace/WorkspaceEditorKit.ts',
  'src/components/transforms.ts',
]);
const hash = (value) => createHash('sha256').update(value).digest('hex');
const inventory = [];
try {
  for (const kit of kits) {
    const path = `src/components/${kit}-kit.tsx`;
    sources.add(path);
    const source = readFileSync(resolve(app, path), 'utf8');
    const file = ts.createSourceFile(path, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    let output = '';
    for (const statement of file.statements) {
      if (!ts.isImportDeclaration(statement)) {
        output += statement.getText(file) + '\n';
        continue;
      }
      if (statement.importClause?.isTypeOnly) continue;
      const specifier = statement.moduleSpecifier.text;
      if (specifier === './indent-kit') {
        output += `import { IndentKit } from './indent.tsx';\n`;
        continue;
      }
      if (specifier.startsWith('@/')) {
        // Stub render components only; no editor transform or plugin implementation is stubbed.
        const bindings = statement.importClause?.namedBindings;
        if (bindings && ts.isNamedImports(bindings))
          for (const entry of bindings.elements)
            if (!entry.isTypeOnly) output += `const ${entry.name.text} = () => null;\n`;
        continue;
      }
      const entry = Bun.resolveSync(specifier, app);
      sources.add(entry.slice(app.length + 1));
      if (specifier.startsWith('@platejs/') || specifier.startsWith('platejs')) {
        const packageName = specifier.startsWith('@')
          ? specifier.split('/').slice(0, 2).join('/')
          : 'platejs';
        const dist = resolve(app, 'node_modules', packageName, 'dist');
        for (const child of readdirSync(dist, { recursive: true }))
          if (String(child).endsWith('.js'))
            sources.add(resolve(dist, String(child)).slice(app.length + 1));
      }
      output +=
        statement
          .getText(file)
          .replace(statement.moduleSpecifier.getText(file), JSON.stringify(entry)) + '\n';
    }
    writeFileSync(resolve(temp, `${kit}.tsx`), output);
    // Exact configured keyboard/input-rule declarations remain reviewable alongside captures.
    const walk = (node) => {
      if (
        ts.isPropertyAssignment(node) &&
        ['inputRules', 'shortcuts'].includes(node.name.getText(file))
      )
        inventory.push({
          source: path,
          kind: node.name.getText(file),
          expression: node.initializer.getText(file),
        });
      ts.forEachChild(node, walk);
    };
    walk(file);
  }
  const consumers = [
    'fixed-toolbar-buttons',
    'floating-toolbar-buttons',
    'turn-into-toolbar-button',
    'turn-into-toolbar-config',
    'mark-toolbar-button',
    'font-size-toolbar-button',
    'font-color-toolbar-button',
    'list-toolbar-button',
    'indent-toolbar-button',
    'align-toolbar-button',
    'line-height-toolbar-button',
    'link-toolbar',
    'link-toolbar-button',
    'link-node',
    'table-toolbar-button',
    'table-node',
    'block-list',
    'slash-node',
    'block-menu',
    'block-menu-items',
    'block-menu-components',
  ];
  for (const name of consumers) {
    const path = `src/components/ui/${name}.tsx`;
    let source;
    try {
      source = readFileSync(resolve(app, path), 'utf8');
    } catch {
      continue;
    }
    sources.add(path);
    const file = ts.createSourceFile(path, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.TSX);
    const calls = new Set();
    const walk = (node) => {
      if (
        ts.isCallExpression(node) &&
        /(?:\.tf\.|^tf\.|setBlockType|insertBlock|toggleList|upsertLink|unwrapLink|use\w*(?:Button|Element))/.test(
          node.expression.getText(file)
        )
      )
        calls.add(node.getText(file));
      ts.forEachChild(node, walk);
    };
    walk(file);
    inventory.push({ source: path, kind: 'consumer', calls: [...calls] });
  }
  sources.add('src/components/ui/slash-catalog.ts');
  const source = readFileSync(resolve(app, 'src/components/transforms.ts'), 'utf8');
  const file = ts.createSourceFile('transforms.ts', source, ts.ScriptTarget.Latest, true);
  const declarations = file.statements.filter(
    (node) =>
      ts.isVariableStatement(node) &&
      node.declarationList.declarations.some((d) =>
        ['LIST_TYPES', 'setBlockType', 'getBlockType', 'insertList', 'insertBlock'].includes(
          d.name.getText(file)
        )
      )
  );
  if (declarations.length !== 5) throw new Error('Production block transform declarations changed');
  const map = file.statements
    .flatMap((node) => (ts.isVariableStatement(node) ? [...node.declarationList.declarations] : []))
    .find((node) => node.name.getText(file) === 'insertBlockMap').initializer;
  const entries = map.properties
    .filter((node) =>
      ['[KEYS.listTodo]', '[KEYS.ol]', '[KEYS.ul]', '[KEYS.table]'].includes(
        node.name.getText(file)
      )
    )
    .map((node) => node.getText(file));
  writeFileSync(
    resolve(temp, 'block.ts'),
    `import { KEYS } from ${JSON.stringify(Bun.resolveSync('platejs', app))};\nimport { setListSiblingNodes } from ${JSON.stringify(Bun.resolveSync('@platejs/list', app))};\nimport { PathApi } from ${JSON.stringify(Bun.resolveSync('platejs', app))};\nimport { insertTable } from ${JSON.stringify(Bun.resolveSync('@platejs/table', app))};\nconst setBlockMap = {};\n${declarations.map((d) => d.getText(file)).join('\n')}\nconst insertBlockMap = {${entries.join(',')}};`
  );
  const runner = readFileSync(
    resolve(root, 'scripts/sciobot-port/native-command-runner.mjs'),
    'utf8'
  ).replace(
    /from '(platejs(?:\/react)?|@platejs\/[^']+)'/g,
    (_, name) => `from ${JSON.stringify(Bun.resolveSync(name, app))}`
  );
  writeFileSync(resolve(temp, 'runner.ts'), runner);
  execFileSync(
    'bun',
    ['--tsconfig-override', resolve(app, 'tsconfig.json'), resolve(temp, 'runner.ts')],
    { cwd: temp, stdio: 'inherit' }
  );
  const result = JSON.parse(readFileSync(resolve(temp, 'results.json'), 'utf8'));
  for (const entry of inventory) {
    const name = entry.source;
    entry.cases = result.cases
      .filter((c) =>
        name.includes('table')
          ? c.group === 'Table'
          : name.includes('link')
            ? c.group === 'Link'
            : name.includes('marks') || name.includes('mark-toolbar')
              ? c.group === 'BasicMarks'
              : name.includes('font')
                ? c.group === 'Font'
                : name.includes('list')
                  ? c.group === 'List'
                  : name.includes('align') || name.includes('indent')
                    ? ['Align', 'List'].includes(c.group)
                    : name.includes('line-height')
                      ? c.group === 'LineHeight'
                      : name.includes('basic-block')
                        ? c.group === 'BasicBlocks'
                        : true
      )
      .map((c) => c.id);
    entry.scope =
      'Document transformations mapped below; UI focus, selection, popover and service calls remain presentation inventory.';
  }
  const artifact = {
    schemaVersion: 1,
    source: {
      revision: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: app, encoding: 'utf8' }).trim(),
      files: Object.fromEntries(
        [...sources].sort().map((path) => [path, hash(readFileSync(resolve(app, path)))])
      ),
      captureHash: hash(
        readFileSync(resolve(root, 'scripts/sciobot-port/capture-native-commands-reference.mjs'))
      ),
      runnerHash: hash(
        readFileSync(resolve(root, 'scripts/sciobot-port/native-command-runner.mjs'))
      ),
      scope:
        'Actual installed Plate transforms with production kit configurations. Render-only imports are inert. Production setBlockType is extracted unchanged; unrelated container dispatch entries are omitted. UI focus/popover/network behavior is inventoried, not executed.',
    },
    inventory,
    ...result,
  };
  const target = resolve(root, 'packages/sciobot/src/__tests__/fixtures/native-commands.json');
  const serialized = await format(JSON.stringify(artifact), {
    ...(await resolveConfig(target)),
    parser: 'json',
  });
  if (process.argv.includes('--check')) {
    if (readFileSync(target, 'utf8') !== serialized)
      throw new Error('Native command capture drift; review source and recapture.');
  } else writeFileSync(target, serialized);
  console.log(`Captured ${artifact.cases.length} native-command reference cases.`);
} finally {
  rmSync(temp, { recursive: true, force: true });
}
