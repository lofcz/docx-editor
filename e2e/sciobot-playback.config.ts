import { defineConfig, devices } from '@playwright/test';
import path from 'node:path';
export default defineConfig({
  testDir: '.',
  testMatch: ['sciobot-playback.spec.ts'],
  workers: 1,
  timeout: 180_000,
  expect: { timeout: 30_000 },
  outputDir: '../screenshots/sciobot-playback-test-results',
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://127.0.0.1:5193',
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
      'SCIOBOT_BROWSER_STABLE=1 bunx vite --config packages/sciobot/test/browser/vite.config.ts --port 5193 --strictPort',
    url: 'http://127.0.0.1:5193/playback.html',
    reuseExistingServer: true,
    timeout: 60_000,
  },
});
