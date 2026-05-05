// @ts-check
//
// W9-F shared-class assertions — pptxjs side.
//
// The renderer now emits a cross-format `oox-*` class alongside each
// format-specific class so manifests can write a single selector that
// works for docxjs/pptxjs/xlsxjs output. This spec verifies the classes
// are present on the expected elements for the basic presentation
// fixture. See src/shared-classes.ts for the concept → class table.
import { test, expect } from '@playwright/test';

test.describe('Shared oox-* classes', () => {
    test('renders with oox-page on slides, oox-paragraph/run on text', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const counts = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/basic/presentation.pptx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await pptx.renderAsync(blob, div);
            const q = (sel) => div.querySelectorAll(sel).length;
            const result = {
                page:      q('.oox-page'),
                paragraph: q('.oox-paragraph'),
                run:       q('.oox-run'),
                // Legacy format-specific class still present (back-compat).
                pptxSlide: q('.pptx-slide'),
            };
            div.remove();
            return result;
        });

        // We don't know the exact fixture composition, so we only assert
        // the shared classes show up at all (and the legacy classes are
        // preserved for back-compat).
        expect(counts.page).toBeGreaterThanOrEqual(1);
        expect(counts.paragraph).toBeGreaterThanOrEqual(1);
        // A presentation with text should have at least one run span.
        expect(counts.run).toBeGreaterThanOrEqual(1);
        // Legacy class still emitted.
        expect(counts.pptxSlide).toBeGreaterThanOrEqual(1);
        // Shared and legacy classes count the same slide elements.
        expect(counts.page).toBe(counts.pptxSlide);
    });
});
