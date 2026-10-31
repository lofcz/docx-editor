import { expect, test } from 'bun:test';
import { fileURLToPath } from 'node:url';
import { applicationLiteralCounts } from './application-literals';

// Sciobot component and storage names belong to the Sciobot registration and its component
// modules. A generic adapter module that names one hard-codes a rule another embedder cannot
// replace, so it reads the configured registry instead.
test('generic adapter modules name no Sciobot component', () => {
  const root = fileURLToPath(new URL('..', import.meta.url));
  expect(applicationLiteralCounts(root)).toEqual({});
});
