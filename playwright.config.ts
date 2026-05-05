import { defineConfig } from '@playwright/test';

// Playwright is configured to drive the system Chrome install (channel: 'chrome')
// because the Playwright-managed Chromium bundle does not yet publish binaries
// for ubuntu26.04 on this host. The tests themselves are plain DOM/fetch and
// don't depend on anything Chrome-specific.
//
// Port 3001 is reserved for pptxjs (docxjs=3000, xlsxjs=3002) so parallel dev
// servers across the three sibling repos don't collide.
export default defineConfig({
  testDir: './tests',
  testMatch: ['**/*.spec.ts', '**/*.spec.js', '**/*.test.ts', '**/*.test.js'],
  fullyParallel: true,
  retries: 0,
  workers: 4,
  reporter: 'list',
  use: {
    baseURL: 'http://localhost:3001',
    trace: 'on-first-retry',
    screenshot: 'only-on-failure',
  },
  projects: [
    {
      name: 'chrome',
      use: {
        browserName: 'chromium',
        channel: 'chrome',
      },
    },
  ],
  webServer: {
    command: 'PORT=3001 node scripts/dev-server.mjs',
    url: 'http://localhost:3001/tests/harness.html',
    reuseExistingServer: !process.env.CI,
    timeout: 30_000,
  },
});
