import { readFileSync } from 'node:fs';
import { dirname, isAbsolute, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const metadata = JSON.parse(readFileSync(resolve(root, 'dist/metafile.json'), 'utf8'));
const inputs = Object.keys(metadata.inputs);
if (!inputs.length) throw new Error('The build has no recorded inputs');
for (const input of inputs) {
  const local = relative(root, isAbsolute(input) ? input : resolve(root, input));
  if (!local.startsWith('src/') || local.includes('__tests__') || /\.(?:test|spec)\./.test(local)) {
    throw new Error(`The private adapter bundled an external runtime or test: ${input}`);
  }
}
let documentEntry = false;
let collaborationEntry = false;
for (const output of Object.values(metadata.outputs)) {
  const entry = output.entryPoint ? relative(root, resolve(root, output.entryPoint)) : null;
  if (entry === 'src/index.ts') {
    documentEntry = true;
    if (!(output.exports ?? []).includes('trackHyperlinkPlan'))
      throw new Error('The document entry does not export the tracked hyperlink executor.');
  }
  if (entry === 'src/collaboration/index.ts') collaborationEntry = true;
  // Bun records external imports on inputs, while its output import lists can be empty.
  const imports = [
    ...(output.imports ?? []),
    ...Object.keys(output.inputs ?? {}).flatMap((input) => metadata.inputs[input]?.imports ?? []),
  ];
  for (const imported of imports) {
    if (entry === 'src/index.ts' && /^react(?:-dom)?(?:\/|$)/.test(imported.path)) {
      throw new Error(`Document tools load the optional React runtime: ${imported.path}`);
    }
    if (entry === 'src/index.ts' && /^yjs(?:\/|$)/.test(imported.path)) {
      throw new Error('Document tools load the optional Yjs runtime.');
    }
    if (
      entry === 'src/collaboration/index.ts' &&
      /^(?:react(?:-dom)?(?:\/|$)|@docx-editor\.dev\/)/.test(imported.path)
    ) {
      throw new Error('The owned attribute protocol imports an editor runtime.');
    }
    if (
      /^(?:@docx-editor\.dev\/(?:pro|editor-api)(?:\/|$)|platejs(?:\/|$)|@platejs\/|@lofcz\/platejs)/.test(
        imported.path
      )
    ) {
      throw new Error(`The private adapter imports an excluded runtime: ${imported.path}`);
    }
  }
}
if (!collaborationEntry) throw new Error('The optional collaboration entry is missing.');
if (!documentEntry) throw new Error('The non-React document entry is missing from the build.');
process.stdout.write(
  `Package boundary passes: ${inputs.length} private source modules; native core stays external.\n`
);
