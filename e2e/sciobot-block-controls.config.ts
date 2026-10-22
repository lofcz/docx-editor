import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-block-controls.spec.ts',
  outputDir: '../screenshots/sciobot-block-controls-results',
});
