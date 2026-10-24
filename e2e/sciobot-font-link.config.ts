import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: ['sciobot-font-link.spec.ts'],
  outputDir: '../screenshots/sciobot-font-link-results',
});
