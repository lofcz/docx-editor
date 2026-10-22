import { defineConfig, devices } from '@playwright/test';
import path from 'node:path';
export default defineConfig({
  testDir: '.',
  testMatch: [
    'sciobot-block-host.spec.ts',
    'sciobot-native-slots.spec.ts',
    'sciobot-block-selection.spec.ts',
    'sciobot-block-input.spec.ts',
    'sciobot-block-clipboard.spec.ts',
    'sciobot-block-measurement.spec.ts',
  ],
  workers: 1,
  timeout: 60_000,
  expect: { timeout: 15_000 },
  outputDir: '../screenshots/sciobot-block-host-test-results',
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://127.0.0.1:5189',
    viewport: { width: 1400, height: 1100 },
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
    launchOptions: {
      executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE ?? '/usr/bin/chromium',
    },
  },
  webServer: {
    cwd: path.resolve(import.meta.dirname, '..'),
    command: 'bunx vite --config packages/sciobot/test/browser/vite.config.ts',
    url: 'http://127.0.0.1:5189',
    reuseExistingServer: true,
    timeout: 30_000,
  },
});
