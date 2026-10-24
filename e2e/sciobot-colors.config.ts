import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-colors.spec.ts',
  outputDir: '../screenshots/sciobot-colors-results',
});
