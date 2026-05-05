import type { Options } from './pptx-preview';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { imageMimeFromPath } from './mime';
import {
	parseSlide,
	parseStaticShapes,
	parseShapesFromSpTree,
	Slide,
	SlideSize,
	Section,
	ShapeLike,
	buildPlaceholderMap,
	emptyPlaceholderMap,
	emptyMasterTextStyles,
	parseMasterTextStyles,
	parseNotesStyle,
	extractHfPlaceholderText,
	emptyHeaderFooterFlags,
	PlaceholderMap,
	MasterTextStyles,
	SlideParseContext,
} from './presentation-parser';
import { LevelStyles } from './text-style';
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
	parseClrMapOvr,
} from './theme';
import { parseSlideBackground, ParsedSlideBackground, BackgroundFill } from './background';
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
const NOTES_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const HANDOUT_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster";

// PresentationML 2010 extensions namespace. Sections (<p14:sectionLst>) are
// PowerPoint 2010+ metadata carried under the standard <p:extLst> extension
// bag in presentation.xml.
const P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main";
// <p:ext uri="…"> value that tags the sectionLst extension. Other extensions
// (e.g. modifyVerifier) live under different URIs in the same extLst.
const SECTION_LST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

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

export interface MasterBundle {
	placeholders: PlaceholderMap;
	textStyles: MasterTextStyles;
	clrMap: ClrMap;
	theme: ThemeColors;
	path: string | null;
	doc: Document | null;
}

// Notes master bundle. Same shape as MasterBundle but carries its own
// `notesStyle` (the notesMaster's <p:notesStyle>, which plays the role
// bodyStyle does for slide masters) alongside the regular textStyles.
// Handout masters use the plain MasterBundle — they have no analogous text
// styles element, but we still want placeholders/theme/clrMap available for
// a future print-handouts pass.
export interface NotesMasterBundle extends MasterBundle {
	notesStyle: LevelStyles;
}

interface LayoutBundle {
	placeholders: PlaceholderMap;
	doc: Document | null;
	// <p:clrMapOvr><a:overrideClrMapping .../></p:clrMapOvr> on the layout,
	// or null when the layout inherits the master's map (either because the
	// element is absent or because it holds <a:masterClrMapping/>).
	clrMapOvr: ClrMap | null;
}

export class Presentation {
	slides: Slide[] = [];
	slideSize: SlideSize = { cx: 9144000, cy: 6858000, type: null, orient: null };
	// rId → resolved URL, merged across all parts (slide/layout/master).
	embedUrls: Map<string, string> = new Map();
	// Populated once at load time from ppt/tableStyles.xml. Null when the
	// package has no such part — renderers should fall back to defaults.
	tableStyles: Map<string, TableStyle> | null = null;
	// <p:presentation firstSlideNum="N"> — the 1-based index to display for
	// the first slide. Defaults to 1 when absent.
	firstSlideNum: number = 1;
	// PowerPoint "Sections" grouping (<p14:sectionLst> under <p:extLst>).
	// Empty when the deck carries no sections — the common case. Hosts can
	// use this to render a section-aware navigator / ToC. Purely metadata;
	// renderer doesn't consume it.
	sections: Section[] = [];
	// Notes master — parsed once from `ppt/notesMasters/notesMaster1.xml` (if
	// declared in the presentation's rels). Carries placeholders + clrMap +
	// theme + the notesStyle text defaults that speaker-notes rendering
	// inherits. Null when the deck has no notesMaster.
	notesMaster: NotesMasterBundle | null = null;
	// Handout master — parsed from `ppt/handoutMasters/handoutMaster1.xml`
	// when present. Currently only loaded (not yet rendered) so future
	// print-handouts support has the chrome/theme context ready.
	handoutMaster: MasterBundle | null = null;

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
			const typeAttr = sldSzEl.getAttribute("type");
			const orientAttr = sldSzEl.getAttribute("orient");
			// orient is a closed ECMA-376 enum (landscape | portrait); anything
			// else is treated as "absent" rather than propagating garbage.
			const orient: 'landscape' | 'portrait' | null =
				orientAttr === 'landscape' || orientAttr === 'portrait' ? orientAttr : null;
			pres.slideSize = {
				cx: Number(sldSzEl.getAttribute("cx")) || pres.slideSize.cx,
				cy: Number(sldSzEl.getAttribute("cy")) || pres.slideSize.cy,
				type: typeAttr || null,
				orient,
			};
		}

		const rels = await pkg.loadRelationships(presPath);
		// Load the notes/handout masters declared directly on the presentation.
		// These parts carry the theme + placeholder + text-style chrome that
		// speaker-notes / print-handouts rendering should inherit. Each is a
		// single optional part — loader helpers return null when absent.
		pres.notesMaster = await loadNotesMaster(pkg, presPath, rels);
		pres.handoutMaster = await loadHandoutMaster(pkg, presPath, rels);

		// Only direct children of <p:sldIdLst> — there may be other <p:sldId>
		// elements nested under <p14:sectionLst> (one per section entry) that
		// use the same local name in the p14 namespace; we don't want those,
		// but the namespace filter below handles it. Still, scope defensively
		// to sldIdLst's direct children to avoid any future ambiguity.
		const sldIdLstEl = presDoc.getElementsByTagNameNS(A_NS.p, "sldIdLst")[0] ?? null;
		const sldIdEls = sldIdLstEl
			? Array.from(sldIdLstEl.children).filter(
				(c) => c.namespaceURI === A_NS.p && c.localName === "sldId",
			)
			: [];
		const slidePaths: string[] = [];
		// p:sldId/@id (the "slide identifier", e.g. 256) → 0-based index in the
		// filtered slidePaths / pres.slides array. Used below to resolve the
		// <p14:sldId id="…"> references inside <p14:section>. Hidden slides
		// that were skipped for rendering are intentionally absent from this
		// map so a section that references them gets those entries omitted.
		const sldIdToIndex = new Map<string, number>();
		for (const el of sldIdEls) {
			// Hidden slide filter: <p:sldId show="0"> is hidden. Per the schema,
			// a missing `show` attr is visible (true). Opt-in override via
			// `showHidden: true` includes them anyway.
			if (!options.showHidden && el.getAttribute("show") === "0") continue;
			const rid = el.getAttributeNS(A_NS.r, "id");
			if (!rid) continue;
			const rel = rels.get(rid);
			if (!rel) continue;
			const sldIdAttr = el.getAttribute("id");
			if (sldIdAttr) sldIdToIndex.set(sldIdAttr, slidePaths.length);
			slidePaths.push(resolveRelTarget(presPath, rel.target));
		}

		// <p14:sectionLst> under <p:extLst>. Each <p14:section> groups a
		// contiguous run of slides (by sldId reference, not by position) with
		// a display name. Purely metadata — rendering is unaffected.
		pres.sections = parseSections(presDoc, sldIdToIndex);
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
			const emptyLayout: LayoutBundle = { placeholders: emptyPlaceholderMap(), doc: null, clrMapOvr: null };

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
					clrMapOvr: layoutDoc ? parseClrMapOvr(layoutDoc) : null,
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

		// Presentation.load walks each slide's layout → master independently
		// (no assumption of a single shared master), so multi-master decks
		// render correctly: every slide pulls the theme/clrMap/textStyles of
		// its own master chain. The caches below are keyed by part path, so
		// distinct masters produce distinct MasterBundles.
		//
		// Caveat: `pres.tableStyles` remains global because the tableStyles
		// part (ppt/tableStyles.xml) is declared at the presentation level,
		// not per-master. Decks that somehow rely on per-master table style
		// variation are not supported — all tables share the same style table.
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

				// Effective clrMap cascade: slide.clrMapOvr ?? layout.clrMapOvr ??
				// master.clrMap. Scheme-color lookups inside this slide's content
				// use this map; master/layout chrome uses its own effective map
				// (layout = layout.clrMapOvr ?? master; master = master.clrMap).
				const slideClrMapOvr = parseClrMapOvr(doc);
				const effectiveClrMap: ClrMap = slideClrMapOvr ?? layout.clrMapOvr ?? master.clrMap;
				const layoutClrMap: ClrMap = layout.clrMapOvr ?? master.clrMap;

				// Shared parse context (clrMap/theme/etc.) that also needs the
				// correct embedUrls for whichever part is being parsed.
				const baseCtx = {
					layout: layout.placeholders,
					master: master.placeholders,
					masterTextStyles: master.textStyles,
					theme: master.theme,
					clrMap: effectiveClrMap,
				};

				// Static chrome from master → layout, parsed with their own embeds
				// so images stored at the master/layout level resolve. Each level
				// parses with the clrMap that applies to it, not the slide's.
				const staticShapes: ReturnType<typeof parseStaticShapes> = [];
				if (master.doc && masterPath) {
					const masterEmbeds = await embedsFor(masterPath);
					const shapes = parseStaticShapes(master.doc, {
						...baseCtx, embedUrls: masterEmbeds, clrMap: master.clrMap,
					});
					rewriteBlipRIds(shapes, masterEmbeds, masterPath, pres.embedUrls);
					staticShapes.push(...shapes);
				}
				if (layout.doc && layoutPath) {
					const layoutEmbeds = await embedsFor(layoutPath);
					const shapes = parseStaticShapes(layout.doc, {
						...baseCtx, embedUrls: layoutEmbeds, clrMap: layoutClrMap,
					});
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

				// Background cascade: slide → layout → master. First concrete hit
				// wins. parseSlideBackground returns 'inherit' when a part carries
				// an explicit <p:bgRef idx=0|1000/> (ECMA-376: "no fill / show
				// through") — that's an intentional declaration meaning "defer to
				// the next level up", distinct from "no <p:bg> declared" (null).
				// Both null and 'inherit' fall through to the next cascade level.
				slide.background = resolveBackgroundCascade(
					parseSlideBackground(doc, effectiveClrMap, master.theme),
					layout.doc ? parseSlideBackground(layout.doc, layoutClrMap, master.theme) : null,
					master.doc ? parseSlideBackground(master.doc, master.clrMap, master.theme) : null,
				);

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

// Reduce the slide → layout → master background cascade down to a concrete
// fill or null. Each input is the result of parseSlideBackground for that
// level: a BackgroundFill (stop here), 'inherit' (explicit skip — look at
// the next level), or null (level said nothing — also look at the next).
function resolveBackgroundCascade(
	slide: ParsedSlideBackground,
	layout: ParsedSlideBackground,
	master: ParsedSlideBackground,
): BackgroundFill | null {
	for (const level of [slide, layout, master]) {
		if (level && level !== 'inherit') return level;
	}
	return null;
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

// Walk <p:extLst> looking for the sectionLst extension and extract the
// section records. `sldIdToIndex` maps each p:sldId/@id to its 0-based
// position in the filtered slide list; sections carrying references to
// hidden / unknown slides simply get those entries dropped.
function parseSections(presDoc: Document, sldIdToIndex: Map<string, number>): Section[] {
	// Direct child search under the root <p:presentation> — extLst can appear
	// at multiple levels in the XML, but sections live specifically in the
	// presentation-level one.
	const root = presDoc.documentElement;
	const extLst = findChild(root, A_NS.p, "extLst");
	if (!extLst) return [];
	const extEls = Array.from(extLst.children).filter(
		(c) => c.namespaceURI === A_NS.p && c.localName === "ext",
	);
	const sectionExt = extEls.find((e) => e.getAttribute("uri") === SECTION_LST_EXT_URI);
	if (!sectionExt) return [];
	const sectionLst = findChild(sectionExt, P14_NS, "sectionLst");
	if (!sectionLst) return [];

	const sections: Section[] = [];
	for (const sec of Array.from(sectionLst.children)) {
		if (sec.namespaceURI !== P14_NS || sec.localName !== "section") continue;
		const id = sec.getAttribute("id") ?? "";
		const name = sec.getAttribute("name") ?? "";
		const sldIdLst = findChild(sec, P14_NS, "sldIdLst");
		const slideIndices: number[] = [];
		if (sldIdLst) {
			for (const sid of Array.from(sldIdLst.children)) {
				if (sid.namespaceURI !== P14_NS || sid.localName !== "sldId") continue;
				const idAttr = sid.getAttribute("id");
				if (!idAttr) continue;
				const idx = sldIdToIndex.get(idAttr);
				if (idx !== undefined) slideIndices.push(idx);
			}
		}
		sections.push({ id, name, slideIndices });
	}
	return sections;
}

function findChild(parent: Element, ns: string, localName: string): Element | null {
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === ns && c.localName === localName) return c;
	}
	return null;
}

// Shared loader for a single notes / handout / slide master part. Resolves
// the part's own theme rel, builds the placeholder map + textStyles, and
// parses the clrMap declared inside the part. Returns null when the part
// itself is missing or unreadable so callers can treat absence as "no such
// master" (distinct from "present but empty").
async function loadMasterBundle(
	pkg: OpenXmlPackage,
	containerPath: string,
	rels: Map<string, import('./open-xml-package').Relationship>,
	relType: string,
): Promise<{ bundle: MasterBundle; doc: Document } | null> {
	let masterPath: string | null = null;
	for (const r of rels.values()) {
		if (r.type === relType) {
			masterPath = resolveRelTarget(containerPath, r.target);
			break;
		}
	}
	if (!masterPath) return null;
	const masterDoc = await pkg.loadXml(masterPath);
	if (!masterDoc) return null;

	let theme = emptyThemeColors();
	const masterRels = await pkg.loadRelationships(masterPath);
	for (const r of masterRels.values()) {
		if (r.type !== THEME_REL_TYPE) continue;
		const themePath = resolveRelTarget(masterPath, r.target);
		const themeDoc = await pkg.loadXml(themePath);
		if (themeDoc) theme = parseTheme(themeDoc);
		break;
	}

	const bundle: MasterBundle = {
		placeholders: buildPlaceholderMap(masterDoc),
		textStyles: parseMasterTextStyles(masterDoc),
		clrMap: parseClrMap(masterDoc),
		theme,
		path: masterPath,
		doc: masterDoc,
	};
	return { bundle, doc: masterDoc };
}

async function loadNotesMaster(
	pkg: OpenXmlPackage,
	presPath: string,
	presRels: Map<string, import('./open-xml-package').Relationship>,
): Promise<NotesMasterBundle | null> {
	const loaded = await loadMasterBundle(pkg, presPath, presRels, NOTES_MASTER_REL_TYPE);
	if (!loaded) return null;
	return { ...loaded.bundle, notesStyle: parseNotesStyle(loaded.doc) };
}

async function loadHandoutMaster(
	pkg: OpenXmlPackage,
	presPath: string,
	presRels: Map<string, import('./open-xml-package').Relationship>,
): Promise<MasterBundle | null> {
	const loaded = await loadMasterBundle(pkg, presPath, presRels, HANDOUT_MASTER_REL_TYPE);
	return loaded ? loaded.bundle : null;
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
