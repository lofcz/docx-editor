import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-line-height.spec.ts',
  outputDir: '../screenshots/sciobot-line-height-results',
});
