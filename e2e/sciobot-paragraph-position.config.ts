import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-paragraph-position.spec.ts',
  outputDir: '../screenshots/sciobot-paragraph-position-results',
});
