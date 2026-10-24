import { expect, test } from 'bun:test';
import { fileURLToPath } from 'node:url';
for (const order of ['fields-first', 'flags-first'])
  test(`field and flag ownership retain both captures and histories when installed ${order}`, () => {
    const fixture = fileURLToPath(
      new URL('../../test/support/review-interception-order.ts', import.meta.url)
    );
    const result = Bun.spawnSync({
      cmd: [process.execPath, 'run', fixture, order],
      stdout: 'pipe',
      stderr: 'pipe',
    });
    expect({ exitCode: result.exitCode, error: new TextDecoder().decode(result.stderr) }).toEqual({
      exitCode: 0,
      error: '',
    });
  });
