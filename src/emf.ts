// EMF / WMF → PNG conversion adapter.
//
// PowerPoint decks routinely embed EMF (Enhanced Metafile) or WMF (Windows
// Metafile) images — typically vector logos exported from Illustrator /
// Visio, or charts pasted from older Office builds. Browsers render neither
// format natively, so without conversion the embedded `<img>` shows a broken
// icon and (worse) leaks the source filename through `alt=`.
//
// The conversion work is delegated to the `emf-converter` npm package,
// which parses the EMF/WMF record stream and replays it onto an in-memory
// Canvas (OffscreenCanvas preferred, HTMLCanvasElement fallback). Output is
// a `data:image/png;base64,…` URL that plugs straight into `<img src=>`.
//
// The library is loaded lazily via dynamic `import()` so it only lands in
// the JS bundle when a deck actually contains an EMF/WMF — decks with only
// PNG/JPEG pay no weight. First-call latency is the cost of fetching the
// chunk + parsing the package (~50-100 KB gzipped). Subsequent calls reuse
// the cached adapter.
//
// Failure modes (all return null, caller should skip registering the URL):
//   - Dynamic import fails (offline, bundler stripped the chunk).
//   - Conversion returns null (malformed EMF, zero-sized bounds, no Canvas
//     API — e.g. headless Node test without a canvas polyfill).
//   - Blob read throws.
//
// PNG is used rather than SVG because `emf-converter` already ships the
// canvas replay and SVG synthesis from GDI records would mean writing a
// second converter. The PNG is rasterised at the EMF's logical bounds so
// at the slide's typical display size it looks effectively the same as
// PowerPoint's native render.

/** Shape of an EMF/WMF → PNG data-URL converter. */
type EmfToDataUrl = (buffer: ArrayBuffer) => Promise<string | null>;

/**
 * Shape of the `emf-converter` module. The adapter accepts anything matching
 * this structural type, so UMD consumers that shim a global
 * (`window.EMFConverter = { convertEmfToDataUrl, convertWmfToDataUrl }`) can
 * participate without reaching for a bundler.
 */
interface EmfConverterModule {
	convertEmfToDataUrl: EmfToDataUrl;
	convertWmfToDataUrl: EmfToDataUrl;
}

// Escape hatch for UMD consumers: pptxjs.setEmfConverter({convertEmfToDataUrl, convertWmfToDataUrl}).
// When set, both adapters are preloaded from it and the dynamic import is skipped.
let emfConverterOverride: EmfConverterModule | null = null;

/**
 * Install an `emf-converter`-compatible module explicitly. Useful for UMD
 * builds where the consumer loads `emf-converter` themselves (e.g. via a
 * separate bundle) and wants to wire it into pptxjs without relying on the
 * bundler-backed `await import()` path. After calling this, any subsequent
 * conversion requests use the provided module directly.
 */
export function setEmfConverter(mod: EmfConverterModule | null): void {
	emfConverterOverride = mod;
	emfAdapter = undefined;
	wmfAdapter = undefined;
}

// Cached adapter. `undefined` = not yet attempted, `null` = load failed
// (give up, don't retry on every image), function = ready to use.
let emfAdapter: EmfToDataUrl | null | undefined = undefined;
let wmfAdapter: EmfToDataUrl | null | undefined = undefined;

/**
 * Lazily resolve the EMF converter. Returns null if the library fails to
 * load for any reason (missing dependency, bundler stripped the chunk,
 * network error in a split-chunk SPA). Callers should treat null as "EMF
 * not supported in this environment" and skip registering a URL for the
 * image, which matches the pre-conversion behaviour.
 */
async function getEmfAdapter(): Promise<EmfToDataUrl | null> {
	if (emfAdapter !== undefined) return emfAdapter;
	if (emfConverterOverride) {
		emfAdapter = (buffer: ArrayBuffer) => emfConverterOverride!.convertEmfToDataUrl(buffer);
		return emfAdapter;
	}
	try {
		const mod = await import('emf-converter');
		emfAdapter = (buffer: ArrayBuffer) => mod.convertEmfToDataUrl(buffer);
	} catch {
		emfAdapter = null;
	}
	return emfAdapter;
}

/** Mirror of `getEmfAdapter` for the WMF entry point on the same package. */
async function getWmfAdapter(): Promise<EmfToDataUrl | null> {
	if (wmfAdapter !== undefined) return wmfAdapter;
	if (emfConverterOverride) {
		wmfAdapter = (buffer: ArrayBuffer) => emfConverterOverride!.convertWmfToDataUrl(buffer);
		return wmfAdapter;
	}
	try {
		const mod = await import('emf-converter');
		wmfAdapter = (buffer: ArrayBuffer) => mod.convertWmfToDataUrl(buffer);
	} catch {
		wmfAdapter = null;
	}
	return wmfAdapter;
}

/**
 * Returns true for the MIME types the EMF/WMF pipeline can rasterise.
 * Callers check this before paying the cost of reading the Blob + loading
 * the converter chunk.
 */
export function isEmfOrWmfMime(mime: string): boolean {
	return mime === "image/x-emf" || mime === "image/emf" ||
		mime === "image/x-wmf" || mime === "image/wmf";
}

/**
 * Convert an EMF/WMF Blob into a `data:image/png;base64,…` URL.
 *
 * Returns null on any failure — the Blob couldn't be read, the converter
 * couldn't be loaded, or the converter itself returned null (invalid input,
 * no canvas available). Caller should treat this as "skip registering a
 * URL for this image" so the rendered `<img>` gets no `src` and the
 * description-leak-via-alt concern stays contained.
 *
 * @param blob - The EMF/WMF binary blob loaded from the PPTX package.
 * @param mime - The MIME the blob was loaded with; determines which
 *   entry point (`convertEmfToDataUrl` vs `convertWmfToDataUrl`) is used.
 */
export async function convertEmfOrWmfBlob(blob: Blob, mime: string): Promise<string | null> {
	const adapter = mime === "image/x-wmf" || mime === "image/wmf"
		? await getWmfAdapter()
		: await getEmfAdapter();
	if (!adapter) return null;
	try {
		const buf = await blob.arrayBuffer();
		return await adapter(buf);
	} catch {
		return null;
	}
}

/**
 * Test hook — resets the memoised adapter so unit tests can re-exercise
 * the lazy-load path. Not part of the public surface.
 *
 * @internal
 */
export function __resetEmfAdaptersForTest(): void {
	emfAdapter = undefined;
	wmfAdapter = undefined;
}
