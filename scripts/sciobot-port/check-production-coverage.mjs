// Development audit only. Parse source text without loading application modules.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync, readdirSync, existsSync, statSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const docs = resolve(root, 'docs/engineering/sciobot-port');
const app = resolve(
  process.env.SCIOBOT_REFERENCE_APP ?? '/run/media/lofcz/ssd_external/GitHub/sciobot-next'
);
const profilePath = resolve(docs, 'evidence/production-coverage.json');
const profile = JSON.parse(readFileSync(profilePath, 'utf8'));
if (profile.schemaVersion !== 1) throw new Error('Unsupported production inventory schema.');
const capture = process.argv.includes('--capture');
if (process.argv.slice(2).some((arg) => arg !== '--capture'))
  throw new Error('Use --capture or no arguments.');
const hashes = {};
const sources = new Map();
function source(path) {
  const absolute = resolve(app, path);
  if (sources.has(absolute)) return sources.get(absolute);
  const bytes = readFileSync(absolute);
  hashes[relative(app, absolute)] = createHash('sha256').update(bytes).digest('hex');
  const file = ts.createSourceFile(
    absolute,
    bytes.toString('utf8'),
    ts.ScriptTarget.Latest,
    true,
    ts.ScriptKind.TSX
  );
  sources.set(absolute, file);
  return file;
}
function locate(specifier, importer) {
  if (!specifier.startsWith('.') && !specifier.startsWith('@/'))
    return relative(app, Bun.resolveSync(specifier, dirname(resolve(app, importer))));
  const stem = specifier.startsWith('@/')
    ? resolve(app, 'src', specifier.slice(2))
    : resolve(app, dirname(importer), specifier);
  const path = [
    stem,
    `${stem}.ts`,
    `${stem}.tsx`,
    resolve(stem, 'index.ts'),
    resolve(stem, 'index.tsx'),
  ].find((candidate) => existsSync(candidate) && statSync(candidate).isFile());
  if (!path) throw new Error(`Cannot resolve ${specifier} from ${importer}`);
  return relative(app, path);
}
function importsOf(path) {
  const names = new Map();
  for (const statement of source(path).statements) {
    if (!ts.isImportDeclaration(statement) || !ts.isStringLiteral(statement.moduleSpecifier))
      continue;
    const bindings = statement.importClause?.namedBindings;
    if (!bindings || !ts.isNamedImports(bindings)) continue;
    for (const entry of bindings.elements)
      names.set(entry.name.text, {
        path: locate(statement.moduleSpecifier.text, path),
        exported: entry.propertyName?.text ?? entry.name.text,
      });
  }
  return names;
}
function initializer(path, name) {
  for (const statement of source(path).statements) {
    if (!ts.isVariableStatement(statement)) continue;
    const declaration = statement.declarationList.declarations.find(
      (entry) => entry.name.getText() === name
    );
    if (declaration?.initializer) return declaration.initializer;
  }
  throw new Error(`Missing ${name} in ${path}`);
}
function definition(path, name) {
  const imported = importsOf(path).get(name);
  const result = imported ?? { path, exported: name };
  source(result.path);
  return result;
}
function components(path) {
  const out = [];
  const visit = (node) => {
    if (
      ts.isCallExpression(node) &&
      ts.isPropertyAccessExpression(node.expression) &&
      node.expression.name.text === 'withComponent' &&
      node.arguments.length === 1 &&
      ts.isIdentifier(node.arguments[0])
    ) {
      const name = node.arguments[0].text;
      out.push({ name, definition: definition(path, name) });
    }
    ts.forEachChild(node, visit);
  };
  visit(source(path));
  return out;
}
function members(path, name) {
  const init = initializer(path, name);
  if (!ts.isArrayLiteralExpression(init))
    throw new Error(`${name} changed composition shape; review its inventory.`);
  return init.elements.map((element) =>
    ts.isSpreadElement(element) ? element.expression : element
  );
}
const leaves = new Set();
for (const name of readdirSync(docs).filter((name) => /^\d\d-.*\.md$/.test(name)))
  for (const match of readFileSync(resolve(docs, name), 'utf8').matchAll(/^## ([A-Z]\d+)$/gm))
    leaves.add(match[1]);
const ledger = new Map(
  [...readFileSync(resolve(docs, 'coverage.md'), 'utf8').matchAll(/^\| `([^`]+)` \| (.+)$/gm)].map(
    (match) => [
      match[1],
      {
        leaves: [...match[2].split('|')[0].matchAll(/\[([A-Z]\d+)\]/g)].map((entry) => entry[1]),
        evidence: match[2].split('|')[2]?.trim(),
      },
    ]
  )
);
const pending = 'Pending production verification';
const workspacePath = 'src/components/workspace/WorkspaceEditorKit.ts';
const workspace = members(workspacePath, 'WorkspaceEditorKit').map((member) => {
  const name = member.getText();
  const feature = profile.features[name];
  const row = ledger.get(name);
  if (
    !feature ||
    !row?.leaves.length ||
    !row.evidence ||
    !feature.behavior ||
    !feature.savedData ||
    !feature.routeSources.length
  )
    throw new Error(
      `Missing source, behavior, stored data, tasks, or verification cell for ${name}`
    );
  for (const id of row.leaves) if (!leaves.has(id)) throw new Error(`Unknown leaf ${id}`);
  for (const path of feature.routeSources) source(path);
  const defined = definition(workspacePath, name);
  return {
    name,
    definition: defined,
    renderers: components(defined.path),
    leaves: row.leaves,
    verification: pending,
  };
});
if (workspace.length !== Object.keys(profile.features).length)
  throw new Error('Unused or duplicate workspace features.');
const staticPath = 'src/components/editor-base-kit.tsx';
const staticKits = members(staticPath, 'BaseEditorKit').map((member) => {
  const name = member.getText();
  const workspaceName = name.replace(/^Base/, '');
  if (name !== 'BaseCommentKit' && !profile.features[workspaceName])
    throw new Error(`Unmapped static kit ${name}`);
  const defined = definition(staticPath, name);
  return {
    name,
    definition: defined,
    renderers: components(defined.path),
    workspaceFeature: name === 'BaseCommentKit' ? null : workspaceName,
    savedData:
      name === 'BaseCommentKit'
        ? 'Stored comment marks and IDs; no live comment composer requirement.'
        : profile.features[workspaceName].savedData,
    leaves: ['I17', 'I21', ...(name === 'BaseCommentKit' ? [] : ledger.get(workspaceName).leaves)],
    verification: pending,
  };
});
const printPath = 'src/components/lesson-plan/lesson-plan-export-kit.tsx';
const printLeaves = {
  BaseLessonTitlePlugin: ['X01'],
  BasePreparationRecapPlugin: ['X02'],
  BaseRecapGoalsPlugin: ['X02'],
  BaseRecapCurriculumPlugin: ['X03'],
  BaseRecapMaterialsPlugin: ['X02'],
  BaseRecapAdaptationsPlugin: ['X02'],
  BaseRecapPhasePlugin: ['X04'],
  BasePhasePlugin: ['X04'],
  BaseActivityPlugin: ['X05'],
  BaseLessonVisualPlugin: ['X08'],
  BaseListeningPlugin: ['X09', 'X10'],
  BaseImagePlugin: ['X26'],
  BaseGalleryPlugin: ['X27'],
};
const print = members(printPath, 'LessonPlanExportKit').map((member) => {
  if (
    !ts.isCallExpression(member) ||
    !ts.isPropertyAccessExpression(member.expression) ||
    member.expression.name.text !== 'withComponent' ||
    member.arguments.length !== 1
  )
    throw new Error('Review changed print renderer composition.');
  const name = member.expression.expression.getText();
  const renderer = member.arguments[0].getText();
  if (!printLeaves[name]) throw new Error(`Unmapped print override ${name}`);
  return {
    name,
    definition: definition(printPath, name),
    renderer,
    rendererDefinition: definition(printPath, renderer),
    leaves: ['I16', 'I17', ...printLeaves[name]],
    verification: pending,
  };
});
for (const path of [
  'src/components/shared-material/SharedLessonDocView.tsx',
  'src/components/workspace/WorkspacePanel.tsx',
  'src/components/workspace/LinkedLessonEditor.tsx',
  'src/lib/workspaceExportSnapshot.ts',
  'src/lib/workspaceExportArchive.ts',
  'src/lib/documentTextExport.ts',
  'src/lib/workspace-lesson-serialize.tsx',
  'src/components/workspace/WorkspaceExportModal.tsx',
  'src/components/chat/renderer/UserMessageWithInlineMarkers.tsx',
  'src/components/chat/chat-input/EdixTextarea.tsx',
])
  source(path);
for (const entry of [...workspace, ...staticKits, ...print])
  for (const id of entry.leaves) if (!leaves.has(id)) throw new Error(`Unknown leaf ${id}`);
const snapshot = {
  applicationRevision: execFileSync('git', ['rev-parse', 'HEAD'], {
    cwd: app,
    encoding: 'utf8',
  }).trim(),
  workspace,
  static: staticKits,
  print,
  files: Object.fromEntries(
    Object.entries(hashes).sort(([left], [right]) => left.localeCompare(right))
  ),
};
if (capture) {
  profile.snapshot = snapshot;
  writeFileSync(profilePath, `${JSON.stringify(profile, null, 2)}\n`);
} else if (JSON.stringify(profile.snapshot) !== JSON.stringify(snapshot))
  throw new Error(
    'Production sources or mappings changed. Review the coverage, then run this command with --capture.'
  );
process.stdout.write(
  `${workspace.length} workspace members, ${staticKits.length} static kits, ${print.length} print overrides; ${Object.keys(hashes).length} source hashes verified.\n`
);
