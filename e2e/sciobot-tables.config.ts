import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-tables.spec.ts',
  outputDir: '../screenshots/sciobot-tables-results',
});
