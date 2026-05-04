// Feature-manifest conformance runner for pptxjs.
//
// Reads every `pptx/*.json` manifest in the sibling corpus
// (../ooxml-reference-corpus/features/pptx/) that carries a
// `render_assertions` block, renders the committed fixture via
// pptxjs's public `renderAsync` entry point, evaluates each render
// assertion, and writes a result JSON mirroring the shape produced by
// `ooxml_validate.conformance.FeatureResult.to_dict()`.
//
// Output path: ../ooxml-validate/conformance/results/pptxjs/pptx/<name>.json
//
// See README.md in the corpus repo for the matrix that consumes these
// JSON files.

import { test, expect } from '@playwright/test';
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  writeFileSync,
} from 'node:fs';
import { join, resolve } from 'node:path';

import {
  evaluateCssSelectorInPage,
  evaluateComputedStyleInPage,
  evaluateVisualSsim,
  type AssertionStatus,
  type AssertionVerdict,
  type RenderAssertion,
} from './evaluator';

// Repo roots. Playwright sets cwd to the project root (where
// playwright.config.ts lives) before loading specs, so we resolve
// sibling checkouts relative to process.cwd() rather than the spec
// file's own location — this avoids needing `import.meta.url`, which
// would force ESM resolution and break pixelmatch/pngjs (CommonJS).
const PPTXJS_ROOT = process.cwd();
const CORPUS_ROOT = resolve(PPTXJS_ROOT, '..', 'ooxml-reference-corpus');
const VALIDATE_ROOT = resolve(PPTXJS_ROOT, '..', 'ooxml-validate');

const FEATURES_DIR = join(CORPUS_ROOT, 'features', 'pptx');
const FIXTURES_DIR = join(CORPUS_ROOT, 'fixtures', 'pptx');
const RESULTS_DIR = join(
  VALIDATE_ROOT,
  'conformance',
  'results',
  'pptxjs',
  'pptx',
);

const LIBRARY_NAME = 'pptxjs';

// Read tool_version out of package.json so result JSONs stay honest
// if the library rev-bumps independently of this runner.
function readToolVersion(): string {
  try {
    const pkg = JSON.parse(
      readFileSync(join(PPTXJS_ROOT, 'package.json'), 'utf-8'),
    );
    return String(pkg.version ?? '');
  } catch {
    return '';
  }
}

interface Manifest {
  id: string;
  title?: string;
  fixtures: { machine: string };
  render_assertions?: RenderAssertion[];
}

interface AssertionResultShape {
  id: string;
  status: AssertionStatus;
  detail: string;
}

interface FeatureResultShape {
  feature_id: string;
  library: string;
  status: AssertionStatus;
  fixture_path: string;
  run_at: string;
  tool_version: string;
  assertions: AssertionResultShape[];
}

function nowIsoUtc(): string {
  // Match Python's `%Y-%m-%dT%H:%M:%SZ` — seconds resolution, no
  // fractional component, trailing Z.
  const d = new Date();
  return d.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function aggregateStatus(
  assertions: AssertionResultShape[],
): AssertionStatus {
  if (assertions.some((a) => a.status === 'error')) return 'error';
  if (assertions.some((a) => a.status === 'fail')) return 'fail';
  return 'pass';
}

/**
 * Discover every pptx feature manifest that carries a
 * `render_assertions` block. Run at test-discovery time so each one
 * becomes its own Playwright test.
 */
function discoverManifests(): { path: string; manifest: Manifest }[] {
  if (!existsSync(FEATURES_DIR)) return [];
  const entries: { path: string; manifest: Manifest }[] = [];
  for (const name of readdirSync(FEATURES_DIR)) {
    if (!name.endsWith('.json')) continue;
    if (name.endsWith('.schema.json')) continue;
    const full = join(FEATURES_DIR, name);
    let parsed: Manifest;
    try {
      parsed = JSON.parse(readFileSync(full, 'utf-8')) as Manifest;
    } catch {
      continue;
    }
    if (!parsed.render_assertions || parsed.render_assertions.length === 0) {
      continue;
    }
    entries.push({ path: full, manifest: parsed });
  }
  return entries;
}

function featureIdToFileName(featureId: string): string {
  // feature_id "pptx/slide-title" → "slide-title.json"
  const parts = featureId.split('/');
  return `${parts[parts.length - 1]}.json`;
}

function fixtureNameFromManifest(manifest: Manifest): string {
  // manifest.fixtures.machine is the logical name, e.g. "pptx/slide-title".
  // Strip the format prefix to get the filename stem.
  const machine = manifest.fixtures.machine;
  const slash = machine.indexOf('/');
  return slash >= 0 ? machine.slice(slash + 1) : machine;
}

function referencePngPath(manifest: Manifest, assertion: RenderAssertion): string {
  if (assertion.reference_png) {
    return resolve(CORPUS_ROOT, assertion.reference_png);
  }
  const stem = fixtureNameFromManifest(manifest);
  return join(CORPUS_ROOT, 'refs', 'pptx', `${stem}-page1.png`);
}

// ---------------------------------------------------------------------------
// Test discovery.
// ---------------------------------------------------------------------------

const corpusPresent = existsSync(FEATURES_DIR);
const discovered = corpusPresent ? discoverManifests() : [];

test.describe('pptxjs render-assertion conformance', () => {
  if (!corpusPresent) {
    test.skip('corpus checkout missing at ../ooxml-reference-corpus', () => {
      // Discovery guard. We don't fail CI just because a developer
      // hasn't cloned the corpus sibling — the matrix simply won't
      // update without it.
    });
    return;
  }

  if (discovered.length === 0) {
    test.skip('no pptx feature manifests with render_assertions yet', () => {});
    return;
  }

  for (const { manifest } of discovered) {
    test(`feature ${manifest.id}`, async ({ page }) => {
      const fixtureStem = fixtureNameFromManifest(manifest);
      const fixtureAbsPath = join(FIXTURES_DIR, `${fixtureStem}.pptx`);
      const fixtureRelPath = `../ooxml-reference-corpus/fixtures/pptx/${fixtureStem}.pptx`;

      const assertions: AssertionResultShape[] = [];

      // Guard: fixture must exist for the run to be meaningful.
      // A manifest-without-fixture is a corpus-authoring gap rather
      // than a renderer defect — skip so the matrix doesn't paint
      // pptxjs red for something it can't fix.
      if (!existsSync(fixtureAbsPath)) {
        test.skip(true, `fixture missing from corpus: ${fixtureAbsPath}`);
        return;
      }

      // Load the harness and render the fixture inside the page.
      const fixtureBytes = readFileSync(fixtureAbsPath);
      const fixtureBase64 = fixtureBytes.toString('base64');

      await page.goto('/tests/harness.html');

      const pageErrors: string[] = [];
      page.on('pageerror', (err) => pageErrors.push(String(err)));

      const renderOk = await page.evaluate(async (b64: string) => {
        const bin = atob(b64);
        const u8 = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
        const blob = new Blob([u8], {
          type:
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        });
        const host = document.getElementById('sink') as HTMLDivElement;
        host.innerHTML = '';
        try {
          // `pptx` is the UMD global from dist/pptx-preview.js.
          // @ts-ignore
          await pptx.renderAsync(blob, host);
          return { ok: true, error: '' };
        } catch (e) {
          return { ok: false, error: String((e as Error).message ?? e) };
        }
      }, fixtureBase64);

      if (!renderOk.ok) {
        assertions.push({
          id: '<render>',
          status: 'error',
          detail: `pptxjs renderAsync threw: ${renderOk.error}`,
        });
      } else if (pageErrors.length > 0) {
        assertions.push({
          id: '<render>',
          status: 'error',
          detail: `page threw uncaught errors: ${pageErrors.join(' | ')}`,
        });
      }

      // Only evaluate render_assertions if rendering did not catastrophically
      // fail. A render error invalidates every downstream DOM / screenshot
      // check.
      if (renderOk.ok && pageErrors.length === 0) {
        for (const ra of manifest.render_assertions ?? []) {
          const verdict = await runOneAssertion(page, manifest, ra);
          assertions.push({ id: ra.id, ...verdict });
        }
      }

      const status = aggregateStatus(assertions);
      writeResult(manifest, fixtureRelPath, assertions);

      // Python-side convention: a fail is recorded signal, not a hard
      // test failure. Only `error` halts the suite.
      expect(status, 'conformance status').not.toBe('error');
    });
  }
});

async function runOneAssertion(
  page: import('@playwright/test').Page,
  manifest: Manifest,
  ra: RenderAssertion,
): Promise<AssertionVerdict> {
  if (ra.kind === 'css_selector') {
    return await page.evaluate(
      // The browser-side evaluator is serialisable; Playwright round-
      // trips it as a Function.prototype.toString() payload.
      (args) => {
        const fn = new Function(
          'assertion',
          `return (${args.src})(assertion);`,
        ) as (a: RenderAssertion) => AssertionVerdict;
        return fn(args.assertion);
      },
      { src: evaluateCssSelectorInPage.toString(), assertion: ra },
    );
  }

  if (ra.kind === 'computed_style') {
    return await page.evaluate(
      (args) => {
        const fn = new Function(
          'assertion',
          `return (${args.src})(assertion);`,
        ) as (a: RenderAssertion) => AssertionVerdict;
        return fn(args.assertion);
      },
      { src: evaluateComputedStyleInPage.toString(), assertion: ra },
    );
  }

  if (ra.kind === 'visual_ssim') {
    const refPath = referencePngPath(manifest, ra);
    return await evaluateVisualSsim(ra, {
      page,
      slideSelector: '.pptx-slide',
      referencePngAbsPath: refPath,
    });
  }

  return {
    status: 'error',
    detail: `Unknown render-assertion kind: ${String((ra as RenderAssertion).kind)}`,
  };
}

function writeResult(
  manifest: Manifest,
  fixtureRelPath: string,
  assertions: AssertionResultShape[],
): void {
  const result: FeatureResultShape = {
    feature_id: manifest.id,
    library: LIBRARY_NAME,
    status: aggregateStatus(assertions),
    fixture_path: fixtureRelPath,
    run_at: nowIsoUtc(),
    tool_version: readToolVersion(),
    assertions,
  };
  mkdirSync(RESULTS_DIR, { recursive: true });
  const outPath = join(RESULTS_DIR, featureIdToFileName(manifest.id));
  writeFileSync(outPath, JSON.stringify(result, null, 2) + '\n', 'utf-8');
}
