import type { Options } from './pptx-preview';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { imageMimeFromPath } from './mime';
import {
	parseSlide,
	parseStaticShapes,
	parseShapesFromSpTree,
	Slide,
	SlideSize,
	ShapeLike,
	buildPlaceholderMap,
	emptyPlaceholderMap,
	emptyMasterTextStyles,
	parseMasterTextStyles,
	extractHfPlaceholderText,
	emptyHeaderFooterFlags,
	PlaceholderMap,
	MasterTextStyles,
	SlideParseContext,
} from './presentation-parser';
import {
	CHART_REL_TYPE,
	findChartFallbackImage,
	findSmartArtDrawingDoc,
	parseSmartArtShapesFromDoc,
} from './graphic-frame';
import {
	ThemeColors,
	ClrMap,
	emptyThemeColors,
	emptyClrMap,
	parseTheme,
	parseClrMap,
} from './theme';
import { parseSlideBackground } from './background';
import { loadTableStyles, TableStyle } from './table-style';
import { loadNotesSlide } from './notes';
import { loadCommentAuthors, loadComments } from './comments';
import { A_NS } from './namespaces';

const LAYOUT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

// Whitelist of URL schemes that may appear in a rendered <a href>. Everything
// else — `javascript:`, `data:`, `file:`, `vbscript:`, custom schemes — is
// dropped on load so a malicious PPTX can't smuggle script execution through
// a hyperlink. `ppaction:` is parsed (for intra-deck jumps) before it reaches
// the anchor and is translated to a `#slide-N` fragment, so it never lands in
// the final href as a bare scheme.
const SAFE_HYPERLINK_SCHEMES = new Set(["http:", "https:", "mailto:", "tel:"]);

export function isSafeHyperlinkHref(href: string): boolean {
	// Fragment-only and relative links are safe (they can't execute code).
	if (href.startsWith("#")) return true;
	// Extract scheme: scheme:[rest]. Must start with a letter and be followed
	// by letters/digits/+-./ per RFC 3986.
	const m = /^([a-zA-Z][a-zA-Z0-9+\-.]*):/.exec(href);
	if (!m) return true; // no scheme → relative URL, safe
	return SAFE_HYPERLINK_SCHEMES.has(m[1].toLowerCase() + ":");
}

interface MasterBundle {
	placeholders: PlaceholderMap;
	textStyles: MasterTextStyles;
	clrMap: ClrMap;
	theme: ThemeColors;
	path: string | null;
	doc: Document | null;
}

interface LayoutBundle {
	placeholders: PlaceholderMap;
	doc: Document | null;
}

export class Presentation {
	slides: Slide[] = [];
	slideSize: SlideSize = { cx: 9144000, cy: 6858000 };
	// rId → resolved URL, merged across all parts (slide/layout/master).
	embedUrls: Map<string, string> = new Map();
	// Populated once at load time from ppt/tableStyles.xml. Null when the
	// package has no such part — renderers should fall back to defaults.
	tableStyles: Map<string, TableStyle> | null = null;
	// <p:presentation firstSlideNum="N"> — the 1-based index to display for
	// the first slide. Defaults to 1 when absent.
	firstSlideNum: number = 1;

	static async load(data: Blob | any, _parser: unknown, options: Options): Promise<Presentation> {
		const pkg = await OpenXmlPackage.load(data, { trimXmlDeclaration: options.trimXmlDeclaration });
		const pres = new Presentation();

		const presPath = "ppt/presentation.xml";
		const presDoc = await pkg.loadXml(presPath);
		if (!presDoc) return pres;

		// Table styles are shared across the whole presentation, so load
		// them once here. An empty map is fine — tables with no matching
		// styleId will just use the per-cell XML.
		pres.tableStyles = await loadTableStyles(pkg);

		const presRoot = presDoc.documentElement;
		const firstSlideNumAttr = presRoot.getAttribute("firstSlideNum");
		if (firstSlideNumAttr) {
			const n = Number(firstSlideNumAttr);
			if (Number.isFinite(n) && n > 0) pres.firstSlideNum = n;
		}

		const sldSzEl = presDoc.getElementsByTagNameNS(A_NS.p, "sldSz")[0];
		if (sldSzEl) {
			pres.slideSize = {
				cx: Number(sldSzEl.getAttribute("cx")) || pres.slideSize.cx,
				cy: Number(sldSzEl.getAttribute("cy")) || pres.slideSize.cy,
			};
		}

		const rels = await pkg.loadRelationships(presPath);
		const sldIdEls = Array.from(presDoc.getElementsByTagNameNS(A_NS.p, "sldId"));
		const slidePaths: string[] = [];
		for (const el of sldIdEls) {
			// Hidden slide filter: <p:sldId show="0"> is hidden. Per the schema,
			// a missing `show` attr is visible (true). Opt-in override via
			// `showHidden: true` includes them anyway.
			if (!options.showHidden && el.getAttribute("show") === "0") continue;
			const rid = el.getAttributeNS(A_NS.r, "id");
			if (!rid) continue;
			const rel = rels.get(rid);
			if (!rel) continue;
			slidePaths.push(resolveRelTarget(presPath, rel.target));
		}
		// Map each slide's package path → its 0-based index, used to translate
		// slide-to-slide relationships (Target="../slides/slide3.xml") to the
		// `#slide-N` anchor fragment the renderer emits on <section>.
		const slidePathToIndex = new Map<string, number>();
		for (let i = 0; i < slidePaths.length; i++) slidePathToIndex.set(slidePaths[i], i);

		const layoutCache = new Map<string, LayoutBundle>();
		const masterCache = new Map<string, MasterBundle>();
		const layoutToMasterPath = new Map<string, string | null>();

		async function getLayoutAndMaster(slidePath: string): Promise<{
			layout: LayoutBundle;
			master: MasterBundle;
			layoutPath: string | null;
			masterPath: string | null;
		}> {
			const emptyMaster: MasterBundle = {
				placeholders: emptyPlaceholderMap(),
				textStyles: emptyMasterTextStyles(),
				clrMap: emptyClrMap(),
				theme: emptyThemeColors(),
				path: null,
				doc: null,
			};
			const emptyLayout: LayoutBundle = { placeholders: emptyPlaceholderMap(), doc: null };

			const slideRels = await pkg.loadRelationships(slidePath);
			let layoutPath: string | null = null;
			for (const r of slideRels.values()) {
				if (r.type === LAYOUT_REL_TYPE) {
					layoutPath = resolveRelTarget(slidePath, r.target);
					break;
				}
			}
			if (!layoutPath) return { layout: emptyLayout, master: emptyMaster, layoutPath: null, masterPath: null };

			let layoutBundle = layoutCache.get(layoutPath);
			if (!layoutBundle) {
				const layoutDoc = await pkg.loadXml(layoutPath);
				layoutBundle = {
					placeholders: layoutDoc ? buildPlaceholderMap(layoutDoc) : emptyPlaceholderMap(),
					doc: layoutDoc,
				};
				layoutCache.set(layoutPath, layoutBundle);

				const layoutRels = await pkg.loadRelationships(layoutPath);
				let masterPath: string | null = null;
				for (const r of layoutRels.values()) {
					if (r.type === MASTER_REL_TYPE) {
						masterPath = resolveRelTarget(layoutPath, r.target);
						break;
					}
				}
				layoutToMasterPath.set(layoutPath, masterPath);
			}

			const masterPath = layoutToMasterPath.get(layoutPath) ?? null;
			let masterBundle = masterPath ? masterCache.get(masterPath) : undefined;
			if (masterPath && !masterBundle) {
				const masterDoc = await pkg.loadXml(masterPath);
				let theme = emptyThemeColors();
				if (masterDoc) {
					const masterRels = await pkg.loadRelationships(masterPath);
					for (const r of masterRels.values()) {
						if (r.type !== THEME_REL_TYPE) continue;
						const themePath = resolveRelTarget(masterPath, r.target);
						const themeDoc = await pkg.loadXml(themePath);
						if (themeDoc) theme = parseTheme(themeDoc);
						break;
					}
				}
				masterBundle = masterDoc
					? {
						placeholders: buildPlaceholderMap(masterDoc),
						textStyles: parseMasterTextStyles(masterDoc),
						clrMap: parseClrMap(masterDoc),
						theme,
						path: masterPath,
						doc: masterDoc,
					}
					: { ...emptyMaster, path: masterPath };
				masterCache.set(masterPath!, masterBundle);
			}

			return {
				layout: layoutBundle,
				master: masterBundle ?? emptyMaster,
				layoutPath,
				masterPath,
			};
		}

		// Author map for <p:cm authorId="…"> lookups. Loaded once up-front —
		// the part either exists (presentation has comments somewhere) or it
		// doesn't, and loadCommentAuthors degrades to an empty map either way.
		const commentAuthors = await loadCommentAuthors(pkg);

		// Cache image URLs across slides — the same embedded image may be
		// referenced by multiple slides via separate rIds, but there's one
		// media entry per unique file.
		const mediaUrlCache = new Map<string, string>();

		async function loadSlideEmbeds(slidePath: string): Promise<Map<string, string>> {
			const urls = new Map<string, string>();
			const slideRels = await pkg.loadRelationships(slidePath);
			for (const rel of slideRels.values()) {
				if (rel.type !== IMAGE_REL_TYPE) continue;
				const mediaPath = resolveRelTarget(slidePath, rel.target);
				let url = mediaUrlCache.get(mediaPath);
				if (!url) {
					const blob = await pkg.loadBlob(mediaPath, imageMimeFromPath(mediaPath));
					if (blob) {
						url = URL.createObjectURL(blob);
						mediaUrlCache.set(mediaPath, url);
					}
				}
				if (url) urls.set(rel.id, url);
			}
			return urls;
		}

		// Per-part embed caches keyed by package path, so layouts/masters with
		// their own images (e.g. a logo on the layout) resolve too.
		async function embedsFor(partPath: string): Promise<Map<string, string>> {
			return loadSlideEmbeds(partPath);
		}

		// Resolve all hyperlink relationships attached to a slide. External
		// targets (http/https/mailto/…) pass straight through after a scheme
		// check. Slide-to-slide rels become `#slide-N` fragments using the
		// presentation's slide order. `ppaction://hlinkshowjump` targets
		// (firstslide / lastslide / nextslide / previousslide / sldjump) are
		// resolved against the current slide's index.
		async function loadSlideHyperlinks(slidePath: string, slideIndex: number): Promise<Map<string, string>> {
			const out = new Map<string, string>();
			const slideRels = await pkg.loadRelationships(slidePath);
			for (const rel of slideRels.values()) {
				if (rel.type === HYPERLINK_REL_TYPE) {
					const target = rel.target;
					if (!target) continue;
					// ppaction://hlinkshowjump?jump=firstslide (and friends). The rel
					// Target carries the ppaction URI verbatim; TargetMode is
					// usually "External" but we don't rely on that — the scheme
					// tells us how to interpret it.
					if (target.startsWith("ppaction://")) {
						const resolved = resolvePpAction(target, slideIndex, slidePaths.length);
						if (resolved) out.set(rel.id, resolved);
						else if (typeof console !== "undefined") console.warn(`pptxjs: unrecognised ppaction hyperlink: ${target}`);
						continue;
					}
					if (!isSafeHyperlinkHref(target)) {
						if (typeof console !== "undefined") console.warn(`pptxjs: dropped hyperlink with unsafe scheme: ${target}`);
						continue;
					}
					out.set(rel.id, target);
				} else if (rel.type === SLIDE_REL_TYPE) {
					// Intra-deck slide-to-slide link. The Target resolves to a
					// slide part path; look it up in the presentation's order.
					const targetPath = resolveRelTarget(slidePath, rel.target);
					const idx = slidePathToIndex.get(targetPath);
					if (idx != null) out.set(rel.id, `#slide-${idx}`);
				}
			}
			return out;
		}

		for (let i = 0; i < slidePaths.length; i++) {
			const path = slidePaths[i];
			// Per-slide isolation: any failure in parse / background / hyperlinks
			// / fallbacks / notes / comments is trapped here so one bad slide
			// doesn't poison the rest of the deck. The slide still takes its
			// slot in `pres.slides` (indices stay stable) but carries a
			// parseError string that the renderer surfaces as a red banner.
			try {
				const doc = await pkg.loadXml(path);
				// A slide referenced from `presentation.xml` whose part is
				// absent (or whose text can't be decoded) raises this — it
				// counts as a parse failure for the purpose of keeping slide
				// indices stable.
				if (!doc) throw new Error(`slide part missing: ${path}`);
				const { layout, master, layoutPath, masterPath } = await getLayoutAndMaster(path);

				// Shared parse context (clrMap/theme/etc.) that also needs the
				// correct embedUrls for whichever part is being parsed.
				const baseCtx = {
					layout: layout.placeholders,
					master: master.placeholders,
					masterTextStyles: master.textStyles,
					theme: master.theme,
					clrMap: master.clrMap,
				};

				// Static chrome from master → layout, parsed with their own embeds
				// so images stored at the master/layout level resolve.
				const staticShapes: ReturnType<typeof parseStaticShapes> = [];
				if (master.doc && masterPath) {
					const masterEmbeds = await embedsFor(masterPath);
					const shapes = parseStaticShapes(master.doc, { ...baseCtx, embedUrls: masterEmbeds });
					rewriteBlipRIds(shapes, masterEmbeds, masterPath, pres.embedUrls);
					staticShapes.push(...shapes);
				}
				if (layout.doc && layoutPath) {
					const layoutEmbeds = await embedsFor(layoutPath);
					const shapes = parseStaticShapes(layout.doc, { ...baseCtx, embedUrls: layoutEmbeds });
					rewriteBlipRIds(shapes, layoutEmbeds, layoutPath, pres.embedUrls);
					staticShapes.push(...shapes);
				}

				const slideEmbeds = await embedsFor(path);
				const slideCtx: SlideParseContext = { ...baseCtx, embedUrls: slideEmbeds };
				const slide = parseSlide(doc, i, slideCtx);
				rewriteBlipRIds(slide.shapes, slideEmbeds, path, pres.embedUrls);

				// Resolve chart / SmartArt fallback shapes now that we have
				// access to the package and the slide's rels.
				const slideRels = await pkg.loadRelationships(path);
				slide.shapes = await resolveGraphicFrameFallbacks(
					pkg, path, slideRels, slide.shapes, slideCtx, mediaUrlCache,
				);

				// Footer/header/datetime text cascade: slide → layout. If the slide
				// doesn't carry the text in its own placeholder, pull it from the
				// layout's corresponding placeholder. Layout placeholders are parsed
				// as part of `staticShapes`.
				if (slide.footerText === null || slide.headerText === null || slide.datetimeText === null) {
					const layoutHfText = extractHfPlaceholderText(staticShapes as ShapeLike[]);
					if (slide.footerText === null) slide.footerText = layoutHfText.ftr;
					if (slide.headerText === null) slide.headerText = layoutHfText.hdr;
					if (slide.datetimeText === null) slide.datetimeText = layoutHfText.dt;
				}

				// Prepend static chrome so slide content renders on top.
				slide.shapes = [...staticShapes, ...slide.shapes];

				// Resolve hyperlink rels for this slide (external URLs + intra-deck
				// jumps). Must be attached after parseSlide since parseSlide returns
				// a fresh empty map.
				slide.hyperlinkUrls = await loadSlideHyperlinks(path, i);

				// Background cascade: slide → layout → master. First hit wins.
				slide.background =
					parseSlideBackground(doc, master.clrMap, master.theme)
					?? (layout.doc ? parseSlideBackground(layout.doc, master.clrMap, master.theme) : null)
					?? (master.doc ? parseSlideBackground(master.doc, master.clrMap, master.theme) : null);

				// Notes + comments — pulled from the slide's own rels. Both are
				// optional parts; loaders return null / [] when absent.
				slide.notes = await loadNotesSlide(pkg, path);
				slide.comments = await loadComments(pkg, path, commentAuthors);
				pres.slides.push(slide);
			} catch (err) {
				if (typeof console !== "undefined") console.warn(`[pptxjs] slide ${i} failed to parse:`, err);
				if (options.onSlideError) {
					try { options.onSlideError(i, err); }
					catch (cbErr) {
						if (typeof console !== "undefined") console.warn(`[pptxjs] onSlideError callback threw:`, cbErr);
					}
				}
				pres.slides.push(makeErrorSlide(i, err));
			}
		}

		return pres;
	}
}

// Build a minimal placeholder slide for a slide that failed to parse. The
// slide still occupies its index in pres.slides so intra-deck `#slide-N`
// anchors stay correct; the renderer checks `parseError` and emits a red
// banner inside an otherwise empty section.
function makeErrorSlide(index: number, err: unknown): Slide {
	const msg = err instanceof Error ? err.message : String(err);
	return {
		index,
		shapes: [],
		background: null,
		hyperlinkUrls: new Map(),
		hf: emptyHeaderFooterFlags(),
		footerText: null,
		headerText: null,
		datetimeText: null,
		notes: null,
		comments: [],
		parseError: msg,
	};
}

// Walk the parsed shapes from a single part (slide/layout/master) and, for
// every BlipFill carrying a rId resolvable via `partEmbeds`, rewrite the
// rId to a globally unique synthetic key and register the corresponding
// URL in `outMap`.
function rewriteBlipRIds(
	shapes: ShapeLike[],
	partEmbeds: Map<string, string>,
	partPath: string,
	outMap: Map<string, string>,
): void {
	for (const s of shapes) {
		if (s.kind !== 'shape') continue;
		if (s.fill && s.fill.kind === 'blip') {
			const url = partEmbeds.get(s.fill.rId);
			if (url) {
				const key = `${partPath}#${s.fill.rId}`;
				outMap.set(key, url);
				s.fill.rId = key;
			}
		}
	}
}

// Walk the slide's shapes after the element-tree parse, filling in chart
// cached images and expanding SmartArt frames into their pre-rendered drawings.
async function resolveGraphicFrameFallbacks(
	pkg: OpenXmlPackage,
	slidePath: string,
	slideRels: Map<string, import('./open-xml-package').Relationship>,
	shapes: ShapeLike[],
	ctx: SlideParseContext,
	mediaUrlCache: Map<string, string>,
): Promise<ShapeLike[]> {
	const out: ShapeLike[] = [];
	for (const s of shapes) {
		if (s.kind === 'chart-fallback') {
			if (s.chartRId) {
				const rel = slideRels.get(s.chartRId);
				if (rel && rel.type === CHART_REL_TYPE) {
					const chartPath = resolveRelTarget(slidePath, rel.target);
					s.src = await findChartFallbackImage(pkg, chartPath, mediaUrlCache);
				}
			}
			out.push(s);
			continue;
		}
		if (s.kind === 'smartart-fallback') {
			if (s.dataRId) {
				const drawingDoc = await findSmartArtDrawingDoc(pkg, slidePath, slideRels, s.dataRId);
				if (drawingDoc) {
					const expanded = parseSmartArtShapesFromDoc(drawingDoc, ctx, parseShapesFromSpTree);
					if (expanded.length > 0) {
						for (const ex of expanded) {
							ex.x += s.x;
							ex.y += s.y;
						}
						out.push(...expanded);
						continue;
					}
				}
			}
			out.push(s);
			continue;
		}
		out.push(s);
	}
	return out;
}

// Translate a `ppaction://hlinkshowjump?jump=...` URI to a `#slide-N` fragment
// relative to the current slide index and total slide count. Returns null for
// unknown or non-mappable actions (e.g. `endshow`) so callers can warn and
// skip.
function resolvePpAction(target: string, slideIndex: number, slideCount: number): string | null {
	const q = target.split("?")[1] ?? "";
	const jumpMatch = /(?:^|&)jump=([^&]+)/.exec(q);
	const action = jumpMatch?.[1];
	switch (action) {
		case "firstslide": return `#slide-0`;
		case "lastslide": return slideCount > 0 ? `#slide-${slideCount - 1}` : null;
		case "nextslide": return slideIndex + 1 < slideCount ? `#slide-${slideIndex + 1}` : null;
		case "previousslide": return slideIndex > 0 ? `#slide-${slideIndex - 1}` : null;
		// `endshow`, `lastslideviewed`, `sldjump` (with no jump= arg), etc. —
		// no sensible HTML mapping, drop silently.
		default: return null;
	}
}
