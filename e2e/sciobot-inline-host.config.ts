import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-inline-host.spec.ts',
  outputDir: '../screenshots/sciobot-inline-host-results',
});
