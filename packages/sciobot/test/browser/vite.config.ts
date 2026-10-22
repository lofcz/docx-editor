import { existsSync, readFileSync } from 'node:fs';
import path from 'node:path';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const root = path.resolve(import.meta.dirname, '../../../..');
const core = path.join(root, 'packages/core');
const manifest = JSON.parse(readFileSync(path.join(core, 'package.json'), 'utf8'));
const aliases = Object.keys(manifest.exports).flatMap((key) => {
  if (key.includes('*') || key.includes('styles') || key.includes('package.json')) return [];
  const subpath = key === '.' ? '' : key.slice(2);
  const stem = path.join(core, 'src', subpath);
  const file = [`${stem}.ts`, path.join(stem, 'index.ts')].find((candidate) =>
    existsSync(candidate)
  );
  if (!file) return [];
  const name = '@docx-editor.dev/core' + (subpath ? `/${subpath}` : '');
  return [
    { find: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`), replacement: file },
  ];
});
// Concurrent checkouts edit watched sources. A stable server avoids reloads during a run.
const stable = process.env.SCIOBOT_BROWSER_STABLE === '1';
export default defineConfig({
  root: import.meta.dirname,
  plugins: [react()],
  resolve: { alias: aliases },
  server: {
    host: '127.0.0.1',
    port: 5189,
    fs: { allow: [root] },
    ...(stable ? { hmr: false, watch: null } : {}),
  },
});
