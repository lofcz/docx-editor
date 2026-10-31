import { expect, test } from 'bun:test';
import { fileURLToPath } from 'node:url';
import baseline from './application-literals.baseline.json';
import { applicationLiteralCounts } from './application-literals';

// Sciobot component and storage names belong to `registry/sciobot*.ts`. This ratchet only
// shrinks: a new literal elsewhere fails, and a removed one must lower the baseline.
test('generic adapter modules name no more Sciobot components than the recorded baseline', () => {
  const root = fileURLToPath(new URL('..', import.meta.url));
  expect(applicationLiteralCounts(root)).toEqual(baseline);
});
