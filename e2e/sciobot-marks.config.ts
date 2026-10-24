import { defineConfig } from '@playwright/test';
import base from './sciobot-block-host.config';
export default defineConfig({
  ...base,
  testMatch: 'sciobot-marks.spec.ts',
  outputDir: '../screenshots/sciobot-marks-results',
});
