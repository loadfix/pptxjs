// Render-assertion evaluators for pptxjs conformance runs.
//
// A render assertion is defined in the corpus manifest schema at
// ../ooxml-reference-corpus/features/manifest.schema.json (see the
// `render_assertion` $def). Three `kind` values are supported:
//
//   - css_selector    — DOM selector count/text check
//   - computed_style  — first match's computed style value
//   - visual_ssim     — screenshot diff vs committed reference PNG
//
// The CSS-selector and computed-style evaluators are written to run
// inside the browser via `page.evaluate`. The visual-SSIM evaluator
// runs Node-side, because reading the reference PNG and executing
// pixelmatch requires filesystem + Node dependencies.
//
// All three return `{status: "pass" | "fail" | "error", detail}` so the
// driver spec can collect verdicts uniformly.

import type { Page } from '@playwright/test';
import { readFileSync, existsSync } from 'node:fs';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';

export type AssertionStatus = 'pass' | 'fail' | 'error';

export interface AssertionVerdict {
  status: AssertionStatus;
  detail: string;
}

export interface RenderAssertion {
  id: string;
  kind: 'css_selector' | 'computed_style' | 'visual_ssim';
  selector?: string;
  must?: 'exist' | 'absent' | 'equal-count' | 'match-text';
  count?: number;
  style_property?: string;
  value?: string;
  min_ssim?: number;
  reference_png?: string;
  description?: string;
}

// ---------------------------------------------------------------------------
// Browser-side evaluators. These are serialisable and passed to
// `page.evaluate`.
// ---------------------------------------------------------------------------

/**
 * Evaluate a css_selector assertion against the live DOM on the page.
 *
 * This function is serialised and executed inside the browser, so it
 * must be self-contained (no outer closures, no imports).
 */
export const evaluateCssSelectorInPage = (
  assertion: RenderAssertion,
): AssertionVerdict => {
  if (!assertion.selector) {
    return {
      status: 'error',
      detail: "css_selector assertion missing 'selector'.",
    };
  }
  const must = assertion.must ?? 'exist';
  const nodes = document.querySelectorAll(assertion.selector);
  const count = nodes.length;

  if (must === 'exist') {
    return count >= 1
      ? { status: 'pass', detail: `${count} node(s) matched.` }
      : { status: 'fail', detail: `No nodes matched selector ${assertion.selector!}.` };
  }

  if (must === 'absent') {
    return count === 0
      ? { status: 'pass', detail: 'No nodes matched (as required).' }
      : { status: 'fail', detail: `${count} node(s) matched but none expected.` };
  }

  if (must === 'equal-count') {
    const expected = assertion.count ?? -1;
    if (expected < 0) {
      return {
        status: 'error',
        detail: "'equal-count' requires a non-negative 'count'.",
      };
    }
    return count === expected
      ? { status: 'pass', detail: `${count} node(s) matched (as required).` }
      : { status: 'fail', detail: `Expected ${expected} node(s), got ${count}.` };
  }

  if (must === 'match-text') {
    if (assertion.value == null) {
      return {
        status: 'error',
        detail: "'match-text' requires a 'value' (regex).",
      };
    }
    if (count === 0) {
      return {
        status: 'fail',
        detail: `No nodes matched selector ${assertion.selector!}; cannot check text.`,
      };
    }
    const actual = nodes[0].textContent ?? '';
    let re: RegExp;
    try {
      re = new RegExp(assertion.value);
    } catch (e) {
      return {
        status: 'error',
        detail: `Invalid regex ${JSON.stringify(assertion.value)}: ${(e as Error).message}`,
      };
    }
    return re.test(actual)
      ? { status: 'pass', detail: `textContent=${JSON.stringify(actual)} matches ${JSON.stringify(assertion.value)}.` }
      : { status: 'fail', detail: `textContent=${JSON.stringify(actual)} does not match ${JSON.stringify(assertion.value)}.` };
  }

  return { status: 'error', detail: `Unknown 'must' mode: ${String(must)}.` };
};

/**
 * Evaluate a computed_style assertion against the live DOM on the page.
 *
 * Runs inside the browser.
 */
export const evaluateComputedStyleInPage = (
  assertion: RenderAssertion,
): AssertionVerdict => {
  if (!assertion.selector) {
    return {
      status: 'error',
      detail: "computed_style assertion missing 'selector'.",
    };
  }
  if (!assertion.style_property) {
    return {
      status: 'error',
      detail: "computed_style assertion missing 'style_property'.",
    };
  }
  if (assertion.value == null) {
    return {
      status: 'error',
      detail: "computed_style assertion missing 'value'.",
    };
  }
  const el = document.querySelector(assertion.selector);
  if (!el) {
    return {
      status: 'fail',
      detail: `No element matched ${assertion.selector}; cannot read computed style.`,
    };
  }
  const actual = getComputedStyle(el as Element).getPropertyValue(
    assertion.style_property,
  );
  let re: RegExp;
  try {
    re = new RegExp(assertion.value);
  } catch (e) {
    return {
      status: 'error',
      detail: `Invalid regex ${JSON.stringify(assertion.value)}: ${(e as Error).message}`,
    };
  }
  return re.test(actual.trim())
    ? {
        status: 'pass',
        detail: `${assertion.style_property}=${JSON.stringify(actual)} matches ${JSON.stringify(assertion.value)}.`,
      }
    : {
        status: 'fail',
        detail: `${assertion.style_property}=${JSON.stringify(actual)} does not match ${JSON.stringify(assertion.value)}.`,
      };
};

// ---------------------------------------------------------------------------
// Node-side evaluator (visual_ssim).
// ---------------------------------------------------------------------------

export interface VisualSsimOptions {
  /** Absolute path to the rendered-slide element's screenshot target. */
  page: Page;
  /** Playwright element locator for the rendered slide (first one). */
  slideSelector: string;
  /** Absolute path to the reference PNG; may not exist. */
  referencePngAbsPath: string;
}

/**
 * Take a screenshot of the first rendered slide and compare it against
 * the reference PNG with pixelmatch.
 *
 * Pixelmatch reports a raw differing-pixel count. We normalise this to
 * a similarity score in [0, 1] (1 - diff/total) and compare against
 * `min_ssim`. This is a pragmatic approximation — true SSIM is a
 * separate algorithm — but is the closest single-dependency option and
 * matches the convention used by the sibling matrix.
 *
 * Dimension mismatch is handled by resizing the screenshot canvas to
 * match the reference PNG via nearest-neighbour blit; anything
 * wildly off-scale will naturally score low.
 */
export async function evaluateVisualSsim(
  assertion: RenderAssertion,
  opts: VisualSsimOptions,
): Promise<AssertionVerdict> {
  const minSsim = assertion.min_ssim ?? 0.95;
  const refPath = opts.referencePngAbsPath;

  if (!existsSync(refPath)) {
    return {
      status: 'error',
      detail: `reference PNG not found at ${refPath}`,
    };
  }

  let reference: PNG;
  try {
    reference = PNG.sync.read(readFileSync(refPath));
  } catch (e) {
    return {
      status: 'error',
      detail: `could not decode reference PNG: ${(e as Error).message}`,
    };
  }

  let shotBuf: Buffer;
  try {
    const slide = opts.page.locator(opts.slideSelector).first();
    if ((await slide.count()) === 0) {
      return {
        status: 'fail',
        detail: `no element matched ${opts.slideSelector}; cannot screenshot.`,
      };
    }
    shotBuf = await slide.screenshot({ type: 'png' });
  } catch (e) {
    return {
      status: 'error',
      detail: `screenshot failed: ${(e as Error).message}`,
    };
  }

  let shot: PNG;
  try {
    shot = PNG.sync.read(shotBuf);
  } catch (e) {
    return {
      status: 'error',
      detail: `could not decode screenshot: ${(e as Error).message}`,
    };
  }

  // pixelmatch requires identical dimensions. If they differ, resize the
  // screenshot into a canvas matching the reference via nearest-neighbour.
  let shotForDiff: PNG = shot;
  if (shot.width !== reference.width || shot.height !== reference.height) {
    shotForDiff = new PNG({ width: reference.width, height: reference.height });
    const sx = shot.width / reference.width;
    const sy = shot.height / reference.height;
    for (let y = 0; y < reference.height; y++) {
      for (let x = 0; x < reference.width; x++) {
        const srcX = Math.min(shot.width - 1, Math.floor(x * sx));
        const srcY = Math.min(shot.height - 1, Math.floor(y * sy));
        const srcIdx = (srcY * shot.width + srcX) * 4;
        const dstIdx = (y * reference.width + x) * 4;
        shotForDiff.data[dstIdx] = shot.data[srcIdx];
        shotForDiff.data[dstIdx + 1] = shot.data[srcIdx + 1];
        shotForDiff.data[dstIdx + 2] = shot.data[srcIdx + 2];
        shotForDiff.data[dstIdx + 3] = shot.data[srcIdx + 3];
      }
    }
  }

  const total = reference.width * reference.height;
  let diff = 0;
  try {
    diff = pixelmatch(
      shotForDiff.data,
      reference.data,
      null,
      reference.width,
      reference.height,
      { threshold: 0.1 },
    );
  } catch (e) {
    return {
      status: 'error',
      detail: `pixelmatch failed: ${(e as Error).message}`,
    };
  }

  const similarity = 1 - diff / total;
  const dimsNote =
    shot.width !== reference.width || shot.height !== reference.height
      ? ` (resized ${shot.width}x${shot.height} → ${reference.width}x${reference.height})`
      : '';

  if (similarity >= minSsim) {
    return {
      status: 'pass',
      detail: `similarity=${similarity.toFixed(4)} >= min_ssim=${minSsim}${dimsNote}`,
    };
  }
  return {
    status: 'fail',
    detail: `similarity=${similarity.toFixed(4)} < min_ssim=${minSsim}${dimsNote}`,
  };
}
