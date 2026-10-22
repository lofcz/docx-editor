import { expect, test } from 'bun:test';
import { realpathSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { createDocxEditor } from '@docx-editor.dev/core';
import { createDocxEditor as createEditorFromSubpath } from '@docx-editor.dev/core/editor';
import manifest from '../../package.json';

test('the private adapter consumes one external native core instance', () => {
  expect(manifest.private).toBe(true);
  expect(manifest.peerDependencies).toEqual({
    '@docx-editor.dev/core': '~2.32.0',
    react: '^18.0.0 || ^19.0.0',
    'react-dom': '^18.0.0 || ^19.0.0',
    yjs: '13.6.32',
  });
  expect(manifest.peerDependenciesMeta).toEqual({
    react: { optional: true },
    'react-dom': { optional: true },
    yjs: { optional: true },
  });
  expect(manifest.exports['./react'].import).toBe('./dist/react-blocks/index.js');
  expect(manifest.exports['./collaboration'].import).toBe('./dist/collaboration/index.js');
  const dependencies = (manifest as { dependencies?: Record<string, string> }).dependencies ?? {};
  expect(dependencies['@docx-editor.dev/core']).toBeUndefined();
  expect(createDocxEditor).toBe(createEditorFromSubpath);
  const corePackage = fileURLToPath(import.meta.resolve('@docx-editor.dev/core/package.json'));
  expect(realpathSync(corePackage)).toBe(
    realpathSync(new URL('../../../core/package.json', import.meta.url))
  );
});
