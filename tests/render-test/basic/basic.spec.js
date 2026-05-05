// @ts-check
import { test, expect } from '@playwright/test';

// Smoke test: load the committed `presentation.pptx` fixture through the
// published `pptx.renderAsync()` entry point and confirm at least one slide
// is produced. There are no pre-existing Karma specs to port in this repo —
// Karma was wired up but never had any `*.spec.js` files. This test
// establishes the Playwright baseline so future specs can follow the same
// harness pattern (identical to loadfix/docxjs).
test.describe('render-test/basic', () => {
    test('renders presentation.pptx without error', async ({ page }) => {
        const errors = [];
        page.on('pageerror', (err) => errors.push(String(err)));

        await page.goto('/tests/harness.html');

        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/basic/presentation.pptx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // `pptx` is the UMD global exposed by dist/pptx-preview.js.
            // @ts-ignore
            const pres = await pptx.renderAsync(blob, div);
            return {
                slideCount: div.querySelectorAll('.pptx-slide').length,
                hasPresentation: !!pres,
                containerHasContent: div.children.length > 0,
            };
        });

        expect(errors, 'page threw no uncaught errors').toEqual([]);
        expect(result.hasPresentation).toBe(true);
        expect(result.containerHasContent).toBe(true);
        expect(result.slideCount).toBeGreaterThan(0);
    });
});
