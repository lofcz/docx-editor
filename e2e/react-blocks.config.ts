import { defineConfig, devices } from '@playwright/test';
export default defineConfig({
  testDir: '.',
  testMatch: 'react-blocks.spec.ts',
  workers: 1,
  timeout: 60_000,
  expect: { timeout: 15_000 },
  outputDir: '../screenshots/react-blocks-test-results',
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://127.0.0.1:5188',
    viewport: { width: 1280, height: 1100 },
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
    launchOptions: { executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE },
  },
  webServer: {
    command: 'bun run --filter docx-editor-example-react-blocks dev',
    url: 'http://127.0.0.1:5188',
    reuseExistingServer: true,
    timeout: 30_000,
  },
});
