import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: ['sciobot-manual-tracking.spec.ts'],
  outputDir: '../screenshots/sciobot-manual-tracking-results',
});
