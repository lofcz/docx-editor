import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: ['sciobot-table-style.spec.ts'],
  outputDir: '../screenshots/sciobot-table-style-results',
});
