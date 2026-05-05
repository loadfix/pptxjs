// @ts-check
import { test, expect } from '@playwright/test';

// Exercises the `responsive: true` render option added in Wave 9 (W9-E).
//
// Reuses the basic `presentation.pptx` fixture from tests/render-test/basic
// since the option is orthogonal to slide content — we assert on the wrapper
// DOM and CSS shape, not on the rendered shapes themselves.

test.describe('responsive mode', () => {
    test('default: no .pptx-slide-container wrapper', async ({ page }) => {
        await page.goto('/tests/harness.html');
        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/basic/presentation.pptx').then((r) => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await pptx.renderAsync(blob, div);
            const out = {
                slideCount: div.querySelectorAll('.pptx-slide').length,
                containerCount: div.querySelectorAll('.pptx-slide-container').length,
            };
            div.remove();
            return out;
        });
        expect(result.slideCount).toBeGreaterThan(0);
        expect(result.containerCount).toBe(0);
    });

    test('responsive: true — each slide is wrapped + transform:scale applied', async ({ page }) => {
        // Constrain the viewport to a phone-ish width so the scale factor is
        // visibly less than 1 (and we can assert on it).
        await page.setViewportSize({ width: 375, height: 812 });
        await page.goto('/tests/harness.html');

        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/basic/presentation.pptx').then((r) => r.blob());
            const div = document.createElement('div');
            // Give the sink a concrete width so the ResizeObserver callback
            // has something to measure against — otherwise clientWidth is 0
            // and the scale factor collapses.
            div.style.width = '350px';
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            await pptx.renderAsync(blob, div, undefined, { responsive: true });
            // ResizeObserver callbacks are async — flush by yielding once.
            await new Promise((r) => requestAnimationFrame(() => r(undefined)));

            const containers = div.querySelectorAll('.pptx-slide-container');
            const firstContainer = /** @type {HTMLElement | null} */ (containers[0]);
            const firstSlide = /** @type {HTMLElement | null} */ (firstContainer?.querySelector('.pptx-slide'));
            const cssText = [...div.querySelectorAll('style')].map((s) => s.textContent ?? '').join('\n');

            const out = {
                containerCount: containers.length,
                slideCount: div.querySelectorAll('.pptx-slide').length,
                containerAspectRatio: firstContainer?.style.aspectRatio ?? '',
                containerWidth: firstContainer?.clientWidth ?? 0,
                transform: firstSlide?.style.transform ?? '',
                cssHasContainerRule: cssText.includes('.pptx-slide-container'),
            };
            div.remove();
            return out;
        });

        expect(result.containerCount).toBeGreaterThan(0);
        expect(result.containerCount).toBe(result.slideCount);
        expect(result.containerAspectRatio).not.toBe('');
        // Transform must be a scale() — any value less than 1 is fine for a
        // phone-width viewport; the exact factor depends on fixture slide size.
        expect(result.transform).toMatch(/^scale\(0?\.\d+\)$/);
        expect(result.cssHasContainerRule).toBe(true);
    });
});
