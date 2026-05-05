// @ts-check
//
// Regression test for the sections parser (Wave 5 A2) and the demo
// sections drawer (Wave 6 A5). The basic fixture carries a spliced
// <p14:sectionLst> with four named sections covering all 20 slides
// (see scripts/make-fixture-basic.py). A previous fixture regeneration
// (commit 54050bb) silently dropped the splice because the regeneration
// was done via pptx-python's normal write path rather than the scripted
// generator, which motivated this guard.
//
// We assert that:
//   1. renderAsync returns a Presentation with the expected sections.
//   2. Each section's slideIndices are 0-based and cover the deck.
//   3. Section names round-trip from the fixture.
import { test, expect } from '@playwright/test';

test.describe('sections', () => {
    test('basic fixture carries a parsed <p14:sectionLst>', async ({ page }) => {
        await page.goto('/tests/harness.html');

        const result = await page.evaluate(async () => {
            const blob = await fetch('/tests/render-test/basic/presentation.pptx').then(r => r.blob());
            const div = document.createElement('div');
            document.body.appendChild(div);
            // @ts-ignore — UMD global
            const pres = await pptx.renderAsync(blob, div);
            const payload = {
                slideCount: pres.slides.length,
                sections: pres.sections.map((s) => ({
                    name: s.name,
                    indices: s.slideIndices,
                })),
            };
            div.remove();
            return payload;
        });

        // Fixture has four sections grouping the 20-slide deck (see
        // scripts/make-fixture-basic.py `_section_defs`).
        expect(result.sections.length).toBe(4);
        const names = result.sections.map((s) => s.name);
        expect(names).toEqual([
            'Introduction',
            'Visuals & Effects',
            'Typography & Text',
            'Tables, Charts & Extras',
        ]);

        // slideIndices are 0-based and every slide belongs to exactly one
        // section. Flattening should yield [0 .. slideCount-1] in order.
        const flat = result.sections.flatMap((s) => s.indices);
        expect(flat).toEqual(
            Array.from({ length: result.slideCount }, (_, i) => i),
        );

        // No section should be empty — the demo drawer's empty-range
        // guard exists but the committed fixture doesn't exercise it.
        for (const s of result.sections) {
            expect(s.indices.length).toBeGreaterThan(0);
        }
    });
});
