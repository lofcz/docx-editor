import { defineConfig, devices } from '@playwright/test';
import path from 'node:path';

export default defineConfig({
  testDir: '.',
  testMatch: ['sciobot-review-restore.spec.ts'],
  workers: 1,
  timeout: 60_000,
  expect: { timeout: 15_000 },
  outputDir: '../screenshots/sciobot-review-restore-test-results',
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://127.0.0.1:5194',
    viewport: { width: 1400, height: 1100 },
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
    launchOptions: {
      executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE ?? '/usr/bin/chromium',
    },
  },
  webServer: {
    cwd: path.resolve(import.meta.dirname, '..'),
    command:
      'SCIOBOT_BROWSER_STABLE=1 bunx vite --config packages/sciobot/test/browser/vite.config.ts --port 5194 --strictPort',
    url: 'http://127.0.0.1:5194/restore.html',
    reuseExistingServer: false,
    timeout: 60_000,
  },
});
