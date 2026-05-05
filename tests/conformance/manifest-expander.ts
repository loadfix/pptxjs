// TypeScript port of `ooxml_validate.conformance.expand_manifest`.
//
// The corpus manifest schema grew a `kind: "parameterised"` variant
// (ooxml-reference-corpus commit 8f4eaf4). One manifest file expands
// into N concrete cases via a Cartesian product over named parameter
// axes. Each expanded case is a fully-literal manifest with
// `assertions` / `render_assertions` materialised from `*_template`
// blocks, `id` + `fixtures.machine` suffixed with the per-axis
// selection ids, and an `_expansion` diagnostic preserved so downstream
// result JSONs can trace back to the source axis values.
//
// No pptx manifests use `kind: "parameterised"` today. The runner
// supports it pre-emptively so a future authoring wave (e.g. a
// "slide-background-color" family) can ship without runner changes.
//
// Semantics mirror the Python reference exactly:
//   - Axes sorted alphabetically before the product. Per-axis values
//     preserve source order. Two authors can reorder a parameter list
//     without invalidating other axes' downstream test ids.
//   - Placeholders match `/{axis.field}/` (lowercase-alpha-numeric-
//     underscore axis + field names). Unknown placeholders pass through
//     unchanged — authoring typos stay visible in the rendered output
//     rather than being silently dropped.
//   - `deepClone` isolates the returned cases from the source manifest
//     so downstream mutation (e.g. enriching a case with library-
//     specific context) is safe.
//
// Exported shape:
//
//   export interface ExpandedManifest { ...manifest fields..., _expansion?; }
//   export function expandManifest(manifest: unknown): ExpandedManifest[];

export interface ExpansionBindings {
  parent_id: string;
  bindings: Record<string, string>;
}

export interface ExpandedManifest {
  id: string;
  kind?: 'literal' | 'parameterised';
  fixtures: { machine: string; office?: string };
  assertions?: unknown[];
  render_assertions?: unknown[];
  _expansion?: ExpansionBindings;
  // All other fields pass through unchanged.
  [key: string]: unknown;
}

interface ParamRecord {
  id: string;
  [key: string]: unknown;
}

type ParameterAxis = ParamRecord[];

/**
 * Deep-clone a JSON-shaped value. Manifests are always JSON so
 * `structuredClone` or a JSON round-trip would work; a hand-rolled
 * walker avoids blowing up on any host-specific values (functions,
 * `undefined` etc.) that the schema forbids anyway.
 */
function deepClone<T>(value: T): T {
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) {
    return value.map((v) => deepClone(v)) as unknown as T;
  }
  const out: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
    out[k] = deepClone(v);
  }
  return out as unknown as T;
}

/**
 * Build the Cartesian product of an array of arrays. Order matches the
 * input order, with the rightmost axis varying fastest — same as
 * `itertools.product` in Python.
 */
function cartesian<T>(axes: T[][]): T[][] {
  if (axes.length === 0) return [[]];
  return axes.reduce<T[][]>(
    (acc, axis) => acc.flatMap((prefix) => axis.map((v) => [...prefix, v])),
    [[]],
  );
}

// Placeholder form: {axis.field}. Both segments are lowercase letters,
// digits, and underscores, starting with a letter. Strict enough to
// avoid eating legitimate `{0}` positional-style tokens while matching
// the Python reference's regex character class exactly.
// Field portion accepts [A-Za-z][A-Za-z0-9_]* to handle camelCase
// OOXML attribute names (e.g. `themeTint`, `numFmtId`). Must match
// the Python reference in ooxml-validate's conformance.py.
const PLACEHOLDER = /\{([a-z][a-z0-9_]*)\.([A-Za-z][A-Za-z0-9_]*)\}/g;

function substituteString(
  text: string,
  bindings: Record<string, ParamRecord>,
): string {
  return text.replace(PLACEHOLDER, (match, axis: string, field: string) => {
    const record = bindings[axis];
    if (!record) return match;
    if (!(field in record)) return match;
    return String(record[field]);
  });
}

function substitute(value: unknown, bindings: Record<string, ParamRecord>): unknown {
  if (typeof value === 'string') return substituteString(value, bindings);
  if (Array.isArray(value)) return value.map((v) => substitute(v, bindings));
  if (value && typeof value === 'object') {
    const out: Record<string, unknown> = {};
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      out[k] = substitute(v, bindings);
    }
    return out;
  }
  return value;
}

/**
 * Expand a feature manifest.
 *
 *   - `kind` omitted or "literal" → `[manifest]` (unchanged, deep-cloned).
 *   - `kind` "parameterised"     → one entry per Cartesian combination.
 *
 * Throws when `kind` is unknown or a parameterised manifest carries no
 * `parameters` block. Matches the Python reference; the runner catches
 * the throw and records the manifest as an error so a single bad
 * authoring file doesn't sink the whole suite.
 */
export function expandManifest(rawManifest: unknown): ExpandedManifest[] {
  if (!rawManifest || typeof rawManifest !== 'object') {
    throw new Error('Manifest must be a JSON object.');
  }
  const manifest = rawManifest as Record<string, unknown>;

  const kind = (manifest.kind as string | undefined) ?? 'literal';

  if (kind === 'literal') {
    return [deepClone(manifest) as ExpandedManifest];
  }
  if (kind !== 'parameterised') {
    throw new Error(`Unknown manifest kind: ${JSON.stringify(kind)}`);
  }

  const parameters = (manifest.parameters ?? {}) as Record<string, ParameterAxis>;
  if (!parameters || Object.keys(parameters).length === 0) {
    throw new Error(
      `Parameterised manifest ${JSON.stringify(manifest.id)} has no parameters.`,
    );
  }

  const axes = Object.keys(parameters).sort();
  const axisValues = axes.map((axis) => parameters[axis]);
  const expanded: ExpandedManifest[] = [];

  for (const combo of cartesian(axisValues)) {
    const bindings: Record<string, ParamRecord> = {};
    for (let i = 0; i < axes.length; i++) {
      bindings[axes[i]] = combo[i];
    }

    const caseObj = deepClone(manifest) as Record<string, unknown>;
    caseObj.kind = 'literal';
    delete caseObj.parameters;

    const suffix = axes.map((axis) => bindings[axis].id).join('--');

    const parentId = String(manifest.id ?? '');
    caseObj.id = `${parentId}--${suffix}`;
    caseObj._expansion = {
      parent_id: parentId,
      bindings: Object.fromEntries(axes.map((axis) => [axis, bindings[axis].id])),
    } satisfies ExpansionBindings;

    // Suffix fixture names so the renderer loads the per-combination
    // fixture, not the shared parent stem. The schema's `fixtures.office`
    // is optional; suffix it only when present.
    const fixtures = (caseObj.fixtures as { machine?: string; office?: string } | undefined) ?? {};
    const machine = fixtures.machine ?? parentId;
    fixtures.machine = `${machine}--${suffix}`;
    if (typeof fixtures.office === 'string') {
      fixtures.office = `${fixtures.office}--${suffix}`;
    }
    caseObj.fixtures = fixtures;

    const assertionsTemplate = caseObj.assertions_template as unknown[] | undefined;
    if (assertionsTemplate !== undefined) {
      caseObj.assertions = assertionsTemplate.map((a) => substitute(a, bindings));
      delete caseObj.assertions_template;
    }
    const renderTemplate = caseObj.render_assertions_template as unknown[] | undefined;
    if (renderTemplate !== undefined) {
      caseObj.render_assertions = renderTemplate.map((a) => substitute(a, bindings));
      delete caseObj.render_assertions_template;
    }

    expanded.push(caseObj as ExpandedManifest);
  }

  return expanded;
}
