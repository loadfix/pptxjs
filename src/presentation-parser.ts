import type { Options } from './pptx-preview';
import { A_NS } from './namespaces';
import { firstChildNS, childrenNS } from './xml-utils';
import type { ChartFallbackShape, SmartArtFallbackShape } from './graphic-frame';
import { URI_TABLE, URI_CHART, URI_DIAGRAM } from './graphic-frame';
import {
	RunStyle,
	ParaStyle,
	LevelStyles,
	BodyProperties,
	emptyRunStyle,
	emptyParaStyle,
	emptyLevelStyles,
	parseRunProps,
	parseParaProps,
	parseLevelStyles,
	parseBodyPr,
	mergeRunStyle,
	mergeParaStyle,
} from './text-style';
import { ThemeColors, ClrMap, resolveSchemeClr, resolveTypeface, resolveColorElement } from './theme';
import { SolidFill, Fill, LineStyle, parseSolidFill, parseFillElement, parseLine } from './fill';
import { ShapeEffects, parseEffectsFromSpPr } from './effects';
import { applyMods } from './color-math';

export interface SlideSize {
	cx: number;
	cy: number;
	// Raw <p:sldSz type="..."> attribute (screen4x3, screen16x9, letter, A4, …).
	// Null when the deck didn't declare one — hosts that care should fall back
	// to inferring from cx/cy. Values are passed through verbatim; pptxjs does
	// not enumerate the ECMA-376 list.
	type: string | null;
	// <p:sldSz orient="..."> — 'landscape' (default) or 'portrait'. Null means
	// the attribute was absent; treat as landscape.
	orient: 'landscape' | 'portrait' | null;
}

// A presentation section from <p14:sectionLst>. Sections group slides for
// navigation / ToC purposes but have no visual effect on rendering.
export interface Section {
	// The section's GUID including braces, e.g. "{ABCD1234-...}" — passed
	// through verbatim from the `id` attribute.
	id: string;
	name: string;
	// 0-based indices into Presentation.slides for each slide referenced by
	// this section's <p14:sldIdLst>. Slides that couldn't be resolved (id not
	// found in the presentation's own sldIdLst) are omitted.
	slideIndices: number[];
}

export interface Slide {
	index: number;
	shapes: ShapeLike[];
	background: import('./background').BackgroundFill | null;
	// Resolved hyperlink URLs for this slide, keyed by relationship id. External
	// URLs (http/https/mailto) are the rel's `Target`. Intra-deck jumps
	// (`ppaction://hlinkshowjump?jump=...` and slide-to-slide relationships) are
	// translated to `#slide-N` fragments. URLs with unsafe schemes
	// (e.g. javascript:) are filtered out during presentation load and not
	// present in this map — the renderer treats missing entries as "no link".
	hyperlinkUrls: Map<string, string>;
	// Header/footer placeholder visibility flags, from <p:cSld><p:hf>.
	// Missing attributes default to true (visible) per the schema.
	hf: HeaderFooterFlags;
	// Resolved footer / header / datetime text. Comes from the slide's own
	// placeholder shapes if present, otherwise from its layout. Used to
	// substitute `ftr` / `hdr` / `dt` field runs at render time.
	footerText: string | null;
	headerText: string | null;
	datetimeText: string | null;
	// Speaker notes from the associated notesSlide part, or null when the
	// slide has none. Populated by Presentation.load after parseSlide runs.
	notes: import('./notes').NotesSlide | null;
	// Review comments on this slide. Always an array — empty when the slide
	// has no comments part.
	comments: import('./comments').Comment[];
	// When the slide could not be parsed, this carries the error's message so
	// the renderer can emit a visible error banner in place of the slide
	// contents. Null on healthy slides. Populated by Presentation.load when
	// parseSlide (or any of the per-slide resolution steps) throws.
	parseError: string | null;
}

export interface HeaderFooterFlags {
	sldNum: boolean;
	hdr: boolean;
	ftr: boolean;
	dt: boolean;
}

export function emptyHeaderFooterFlags(): HeaderFooterFlags {
	return { sldNum: true, hdr: true, ftr: true, dt: true };
}

// Parse <p:cSld><p:hf ... /> — attributes `sldNum`, `hdr`, `ftr`, `dt`
// are "1"/"0"; a missing attribute means visible.
export function parseHeaderFooterFlags(cSld: Element | null): HeaderFooterFlags {
	const out = emptyHeaderFooterFlags();
	if (!cSld) return out;
	const hf = firstChildNS(cSld, A_NS.p, "hf");
	if (!hf) return out;
	const read = (name: string) => {
		const v = hf.getAttribute(name);
		return v === null ? true : v !== "0";
	};
	out.sldNum = read("sldNum");
	out.hdr = read("hdr");
	out.ftr = read("ftr");
	out.dt = read("dt");
	return out;
}

export type ShapeLike = Shape | PicShape | TableShape | ChartFallbackShape | SmartArtFallbackShape;

// Non-visual descriptive metadata extracted from <p:cNvPr>. Shared across
// every shape kind (sp, pic, cxnSp, graphicFrame, grpSp).
export interface NonVisualProps {
	name: string | null;
	title: string | null;
	descr: string | null;
}

export function emptyNonVisualProps(): NonVisualProps {
	return { name: null, title: null, descr: null };
}

// Parse the <p:cNvPr> child of any <p:nvSpPr>/<p:nvPicPr>/<p:nvCxnSpPr>
// etc. The `nvSpPrLike` argument is the wrapper element containing cNvPr —
// callers pass e.g. firstChildNS(sp, A_NS.p, "nvSpPr").
export function parseNonVisualProps(nvSpPrLike: Element | null): NonVisualProps {
	const out = emptyNonVisualProps();
	if (!nvSpPrLike) return out;
	const cNvPr = firstChildNS(nvSpPrLike, A_NS.p, "cNvPr");
	if (!cNvPr) return out;
	out.name = cNvPr.getAttribute("name") || null;
	out.title = cNvPr.getAttribute("title") || null;
	out.descr = cNvPr.getAttribute("descr") || null;
	return out;
}

// Find the <a:hlinkClick> child of a <p:cNvPr> element, if any.
function hyperlinkFromCNvPr(nvSpPrLike: Element | null): string | null {
	if (!nvSpPrLike) return null;
	const cNvPr = firstChildNS(nvSpPrLike, A_NS.p, "cNvPr");
	if (!cNvPr) return null;
	const hlinkClick = firstChildNS(cNvPr, A_NS.a, "hlinkClick");
	if (!hlinkClick) return null;
	return hlinkClick.getAttributeNS(A_NS.r, "id") || null;
}

// Mirrors <a:tblPr> boolean flags that toggle band application.
export interface TableFlags {
	firstRow: boolean;
	firstCol: boolean;
	lastRow: boolean;
	lastCol: boolean;
	bandRow: boolean;
	bandCol: boolean;
}

export function emptyTableFlags(): TableFlags {
	return { firstRow: false, firstCol: false, lastRow: false, lastCol: false, bandRow: false, bandCol: false };
}

export interface TableShape {
	kind: 'table';
	x: number;
	y: number;
	cx: number;
	cy: number;
	colWidthsEmu: number[];
	rows: TableRow[];
	// The UUID reference into ppt/tableStyles.xml. Null when <a:tableStyleId>
	// wasn't present. The renderer looks this up in the shared styles map.
	styleId: string | null;
	tableFlags: TableFlags;
	// Shared shape metadata.
	name: string | null;
	title: string | null;
	alt: string | null;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
}

export interface TableRow {
	heightEmu: number;
	cells: TableCell[];
}

// Per-side borders carried by <a:tcPr>. Each LineStyle is the direct
// result of parseLine on the corresponding <a:lnL|lnR|lnT|lnB|...>.
export interface TableCellBorders {
	l: LineStyle | null;
	r: LineStyle | null;
	t: LineStyle | null;
	b: LineStyle | null;
	tlbr: LineStyle | null;
	blTr: LineStyle | null;
}

export interface TableCell {
	paragraphs: Paragraph[];
	// Any fill kind from <a:tcPr> (not just solid). Null when the cell
	// defers to the table-style band.
	fill: Fill | null;
	// Column span > 1 (from `gridSpan`); merged continuation cells (`hMerge`) are
	// dropped during parse. Similarly `rowSpan` / `vMerge` handling is minimal.
	gridSpan: number;
	rowSpan: number;
	// Continuation cells that the renderer should skip.
	hMerge: boolean;
	vMerge: boolean;
	// Per-side padding (marL/marR/marT/marB) in EMU. Null means inherit
	// defaults (the renderer applies ~3.6pt = ~91440 EMU, matching
	// PowerPoint's default).
	insetsEmu: { l: number; r: number; t: number; b: number } | null;
	// <a:tcPr anchor="t|ctr|b"> — vertical text alignment inside the cell.
	anchor: 't' | 'ctr' | 'b' | null;
	borders: TableCellBorders;
}

export interface Shape {
	kind: 'shape';
	x: number;
	y: number;
	cx: number;
	cy: number;
	fill: Fill | null;
	line: LineStyle | null;
	paragraphs: Paragraph[];
	// Custom geometry (<a:custGeom>), if any. Null for preset/no geometry.
	// When present, the renderer produces an SVG path instead of a plain box.
	custGeom: CustGeom | null;
	// Preset geometry (<a:prstGeom prst="…">), if any. custGeom takes precedence
	// when both are set. `avLst` captures the shape's <a:gd> adjust values
	// (e.g. adj1 → 15000) used to parameterise the preset silhouette.
	presetGeom: { name: string; avLst: Map<string, number> } | null;
	// Parsed <a:effectLst>. Null when no effects are present.
	effects: ShapeEffects | null;
	// Accessibility / descriptive metadata. Null when the PPTX didn't set it.
	name: string | null;
	title: string | null;
	alt: string | null;
	// Click-hyperlink rId from <p:cNvPr><a:hlinkClick>.
	hyperlinkRId: string | null;
	// Rotation in 60000ths of a degree. flipH/flipV are the xfrm boolean flags.
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
	// Normalized placeholder type from <p:ph type="..."> (e.g. "sldNum",
	// "ftr", "hdr", "dt", "title", "body"). Null for non-placeholder shapes.
	// Consumed by the renderer so it can skip sldNum/ftr/hdr/dt placeholders
	// when the slide's <p:hf> flag disables them.
	phType: string | null;
	// Parsed <a:bodyPr> from <p:txBody>. Null when the shape has no text body
	// or the element was absent. Drives text-frame insets, vertical anchor,
	// wrap, column count, writing-mode, and autofit scaling.
	bodyPr: BodyProperties | null;
}

// Minimal <a:custGeom> representation — a concatenation of every
// <a:path> inside <a:pathLst>, each as a subpath in the resulting SVG `d`
// string. The local coordinate dimensions (`pathW` / `pathH`) come from
// the first path's `w` / `h` attributes (PowerPoint rarely mixes sizes
// across paths in a single custGeom, so we pick the first and use that
// as the viewBox). `closed` is true when at least one subpath ended
// with <a:close>.
export interface CustGeom {
	pathW: number;
	pathH: number;
	// SVG `d` attribute built during parse. Expressed in the path's local
	// coordinate system (0..pathW, 0..pathH) so the renderer can emit it
	// inside an SVG viewBox matching those dimensions.
	d: string;
	closed: boolean;
	// <a:path fill="…"> — 'none' (no fill), 'norm' (use shape fill),
	// 'darken' / 'lighten' (modify shape fill, currently treated as norm).
	// Taken from the first subpath; multi-path shapes that mix fill modes
	// are rare and not fully round-tripped here.
	fillMode: 'none' | 'norm' | 'darken' | 'lighten';
}

export interface PicShape {
	kind: 'pic';
	x: number;
	y: number;
	cx: number;
	cy: number;
	// Pre-resolved URL (blob: or data:) for the image. Null if the embed
	// relationship couldn't be resolved — render as an empty box.
	src: string | null;
	alt: string;
	name: string | null;
	title: string | null;
	hyperlinkRId: string | null;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
	// Source-rect crop from <a:srcRect>. Each value is in OOXML per-mille
	// (0..100000) and expresses the fraction of the source image clipped from
	// that edge. Missing attributes default to 0. Null means no crop.
	srcRectPermille: SrcRect | null;
	// <a:stretch> sets this true (the default when neither stretch nor tile
	// is present); <a:tile> sets this false and populates `tile`.
	stretch: boolean;
	// <a:tile> metadata — scale/offset values in OOXML per-mille units. Null
	// unless the blipFill explicitly used <a:tile>.
	tile: TileInfo | null;
	// Fixed-alpha override from <a:blip><a:alphaModFix amt="N"/>. N is in
	// OOXML per-mille (0..100000). Null when absent.
	alphaPermille: number | null;
	// <a:lum bright="N" contrast="N"/>. Both in per-mille; can be negative.
	lumBrightPermille: number | null;
	lumContrastPermille: number | null;
	// <a:grayscl/> presence flag.
	grayscale: boolean;
	// <a:biLevel thresh="N"/> — threshold in per-mille. Null when absent.
	biLevelPermille: number | null;
	// <a:duotone> with two color children — resolved hex colors.
	duotone: [string, string] | null;
}

export interface SrcRect {
	l: number;
	t: number;
	r: number;
	b: number;
}

export interface TileInfo {
	sxPermille: number;
	syPermille: number;
	txPermille: number;
	tyPermille: number;
}

export interface Paragraph {
	level: number;
	style: ParaStyle;
	runs: Run[];
}

// Runs now come in three flavours:
//   - TextRun: ordinary <a:r><a:t>text</a:t></a:r>.
//   - BreakRun: <a:br/>, a hard line break within a paragraph.
//   - FieldRun: <a:fld type="slidenum|..."/>, auto-populated text (slide
//     number, footer, date). `fallbackText` is the embedded <a:t> content
//     used when the field can't be resolved by the renderer.
export type Run = TextRun | BreakRun | FieldRun;

export interface TextRun {
	kind: 'text';
	text: string;
	style: RunStyle;
}

export interface BreakRun {
	kind: 'break';
}

export interface FieldRun {
	kind: 'field';
	fieldType: string;
	fallbackText: string;
	style: RunStyle;
}

export interface PlaceholderFrame {
	x: number;
	y: number;
	cx: number;
	cy: number;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
}

export interface PlaceholderInfo {
	frame: PlaceholderFrame | null;
	// Per-level run defaults contributed by this placeholder's <a:lstStyle>.
	levelStyles: LevelStyles;
}

// Master-level fallback styles, keyed by placeholder "category".
// Per ECMA-376: titleStyle → title/ctrTitle, bodyStyle → body/subTitle, otherStyle → everything else.
export interface MasterTextStyles {
	title: LevelStyles;
	body: LevelStyles;
	other: LevelStyles;
}

export interface PlaceholderMap {
	byKey: Map<string, PlaceholderInfo>;
	byType: Map<string, PlaceholderInfo>;
}

export function emptyPlaceholderMap(): PlaceholderMap {
	return { byKey: new Map(), byType: new Map() };
}

export function emptyMasterTextStyles(): MasterTextStyles {
	return { title: emptyLevelStyles(), body: emptyLevelStyles(), other: emptyLevelStyles() };
}

function normalizePhType(t: string | null): string {
	if (!t) return "body";
	if (t === "ctrTitle") return "title";
	return t;
}

// Maps a placeholder type to one of the three master txStyles buckets.
function phTypeToStyleBucket(type: string): keyof MasterTextStyles {
	if (type === "title") return "title";
	if (type === "body" || type === "subTitle") return "body";
	return "other";
}

export function buildPlaceholderMap(doc: Document): PlaceholderMap {
	const map = emptyPlaceholderMap();
	const sps = Array.from(doc.getElementsByTagNameNS(A_NS.p, "sp"));
	for (const sp of sps) {
		const nvSpPr = firstChildNS(sp, A_NS.p, "nvSpPr");
		const nvPr = nvSpPr && firstChildNS(nvSpPr, A_NS.p, "nvPr");
		const ph = nvPr && firstChildNS(nvPr, A_NS.p, "ph");
		if (!ph) continue;
		const type = normalizePhType(ph.getAttribute("type"));
		const idx = ph.getAttribute("idx") ?? "0";

		const spPr = firstChildNS(sp, A_NS.p, "spPr");
		const frame = readXfrm(spPr);

		const txBody = firstChildNS(sp, A_NS.p, "txBody");
		const lstStyle = txBody ? firstChildNS(txBody, A_NS.a, "lstStyle") : null;
		const levelStyles = parseLevelStyles(lstStyle);

		const info: PlaceholderInfo = { frame, levelStyles };

		map.byKey.set(`${type}:${idx}`, info);
		if (!map.byType.has(type)) map.byType.set(type, info);
	}
	return map;
}

// Parse the master's <p:txStyles> (siblings of <p:cSld>).
export function parseMasterTextStyles(masterDoc: Document): MasterTextStyles {
	const root = masterDoc.documentElement;
	const txStyles = firstChildNS(root, A_NS.p, "txStyles");
	if (!txStyles) return emptyMasterTextStyles();
	return {
		title: parseLevelStyles(firstChildNS(txStyles, A_NS.p, "titleStyle")),
		body: parseLevelStyles(firstChildNS(txStyles, A_NS.p, "bodyStyle")),
		other: parseLevelStyles(firstChildNS(txStyles, A_NS.p, "otherStyle")),
	};
}

export class PresentationParser {
	constructor(public options: Options) {}
}

export interface SlideParseContext {
	layout: PlaceholderMap;
	master: PlaceholderMap;
	masterTextStyles: MasterTextStyles;
	theme: ThemeColors;
	clrMap: ClrMap;
	// rId → resolved URL (blob: or data:) for images already loaded from
	// this slide's relationships. Built before parseSlide runs.
	embedUrls: Map<string, string>;
}

export function parseSlide(doc: Document, index: number, ctx: SlideParseContext): Slide {
	const shapes: ShapeLike[] = [];
	const cSld = firstChildNS(doc.documentElement, A_NS.p, "cSld");
	const spTree = firstChildNS(cSld, A_NS.p, "spTree");
	if (spTree) walkSpTree(spTree, shapes, ctx);
	const hf = parseHeaderFooterFlags(cSld);
	// Pull placeholder text for ftr/hdr/dt so FieldRuns can substitute even
	// when the slide itself has no dedicated placeholder shape (the layout's
	// text is the cascade fallback).
	const phText = extractHfPlaceholderText(shapes);
	return {
		index,
		shapes,
		background: null,
		hyperlinkUrls: new Map(),
		hf,
		footerText: phText.ftr,
		headerText: phText.hdr,
		datetimeText: phText.dt,
		notes: null,
		comments: [],
		parseError: null,
	};
}

// Walk a shape list and pull plain text from sldNum/ftr/hdr/dt placeholders.
// Used to resolve `ftr` / `hdr` / `dt` field substitution — PowerPoint stores
// the actual footer/header string inside a placeholder shape whose ph.type
// marks its role.
export function extractHfPlaceholderText(shapes: ShapeLike[]): {
	ftr: string | null; hdr: string | null; dt: string | null;
} {
	let ftr: string | null = null;
	let hdr: string | null = null;
	let dt: string | null = null;
	for (const s of shapes) {
		if (s.kind !== 'shape' || !s.phType) continue;
		const txt = paragraphsPlainText(s.paragraphs);
		if (!txt) continue;
		if (s.phType === 'ftr' && ftr === null) ftr = txt;
		else if (s.phType === 'hdr' && hdr === null) hdr = txt;
		else if (s.phType === 'dt' && dt === null) dt = txt;
	}
	return { ftr, hdr, dt };
}

function paragraphsPlainText(paras: Paragraph[]): string {
	const lines: string[] = [];
	for (const p of paras) {
		let s = "";
		for (const r of p.runs) {
			if (r.kind === 'text') s += r.text;
			else if (r.kind === 'break') s += "\n";
			// field runs inside a placeholder's own text are ignored — they'd
			// recurse back into resolution.
		}
		if (s) lines.push(s);
	}
	return lines.join("\n");
}

// Collect non-placeholder shapes from a layout or master — static decoration
// (lines, logos, background images) that should render on every slide using
// the layout/master. Placeholder shapes (those with <p:ph>) are excluded
// because they're either materialised by the slide's own <p:sp> (when the
// user populated them) or deliberately inherited via frame lookup.
//
// `showMasterPh` should be false for layouts (the layout's ph children are
// pure templates) and true for masters when processed separately.
export function parseStaticShapes(doc: Document, ctx: SlideParseContext): ShapeLike[] {
	const out: ShapeLike[] = [];
	const spTree = firstChildNS(firstChildNS(doc.documentElement, A_NS.p, "cSld"), A_NS.p, "spTree");
	if (!spTree) return out;
	walkSpTreeStatic(spTree, out, ctx);
	return out;
}

function walkSpTreeStatic(container: Element, out: ShapeLike[], ctx: SlideParseContext, t: Transform = IDENTITY): void {
	for (const child of Array.from(container.children)) {
		if (child.namespaceURI !== A_NS.p) continue;
		// Per-shape isolation: one mangled shape shouldn't abort parsing of
		// its siblings. Drop the shape + warn; grpSp failures drop the whole
		// group (but not surrounding shapes).
		try {
			if (child.localName === "sp" || child.localName === "cxnSp") {
				if (isPlaceholder(child)) continue;
				const s = parseShape(child, ctx);
				if (s) { applyTransform(s, t); out.push(s); }
			} else if (child.localName === "pic") {
				if (isPlaceholder(child)) continue;
				const p = parsePic(child, ctx);
				if (p) { applyTransform(p, t); out.push(p); }
			} else if (child.localName === "grpSp") {
				const grpT = composeTransforms(t, transformFromGrpSp(child));
				walkSpTreeStatic(child, out, ctx, grpT);
			} else if (child.localName === "graphicFrame") {
				const gfShape = parseGraphicFrame(child, ctx);
				if (gfShape) { applyTransform(gfShape, t); out.push(gfShape); }
			}
		} catch (err) {
			if (typeof console !== "undefined") console.warn(`[pptxjs] shape <${child.localName}> failed to parse, dropping:`, err);
		}
	}
}

function isPlaceholder(shape: Element): boolean {
	const nvSpPr = firstChildNS(shape, A_NS.p, "nvSpPr")
		?? firstChildNS(shape, A_NS.p, "nvCxnSpPr")
		?? firstChildNS(shape, A_NS.p, "nvPicPr");
	if (!nvSpPr) return false;
	const nvPr = firstChildNS(nvSpPr, A_NS.p, "nvPr");
	return !!(nvPr && firstChildNS(nvPr, A_NS.p, "ph"));
}

// Affine scale+translate applied in slide-global EMU coordinates. Identity is
// scale=1, offset=0 — i.e. no transform.
interface Transform {
	scaleX: number;
	scaleY: number;
	offsetX: number;
	offsetY: number;
}

const IDENTITY: Transform = { scaleX: 1, scaleY: 1, offsetX: 0, offsetY: 0 };

// Compose outer ∘ inner — apply outer to the result of inner.
function composeTransforms(outer: Transform, inner: Transform): Transform {
	return {
		scaleX: outer.scaleX * inner.scaleX,
		scaleY: outer.scaleY * inner.scaleY,
		offsetX: outer.offsetX + outer.scaleX * inner.offsetX,
		offsetY: outer.offsetY + outer.scaleY * inner.offsetY,
	};
}

// Build a Transform for a <p:grpSp>'s own <a:xfrm>. `off`/`ext` describe the
// group's position in parent coords; `chOff`/`chExt` describe the child-local
// coordinate frame. Mapping: parent = off + (child - chOff) * (ext/chExt).
function transformFromGrpSp(grpSp: Element): Transform {
	const grpSpPr = firstChildNS(grpSp, A_NS.p, "grpSpPr");
	const xfrm = grpSpPr && firstChildNS(grpSpPr, A_NS.a, "xfrm");
	if (!xfrm) return IDENTITY;
	const off = firstChildNS(xfrm, A_NS.a, "off");
	const ext = firstChildNS(xfrm, A_NS.a, "ext");
	const chOff = firstChildNS(xfrm, A_NS.a, "chOff");
	const chExt = firstChildNS(xfrm, A_NS.a, "chExt");
	if (!off || !ext || !chOff || !chExt) return IDENTITY;
	const ox = Number(off.getAttribute("x")) || 0;
	const oy = Number(off.getAttribute("y")) || 0;
	const ecx = Number(ext.getAttribute("cx")) || 0;
	const ecy = Number(ext.getAttribute("cy")) || 0;
	const chox = Number(chOff.getAttribute("x")) || 0;
	const choy = Number(chOff.getAttribute("y")) || 0;
	const checx = Number(chExt.getAttribute("cx")) || 0;
	const checy = Number(chExt.getAttribute("cy")) || 0;
	// Guard against degenerate child frames (divide-by-zero).
	const scaleX = checx === 0 ? 1 : ecx / checx;
	const scaleY = checy === 0 ? 1 : ecy / checy;
	return {
		scaleX,
		scaleY,
		offsetX: ox - chox * scaleX,
		offsetY: oy - choy * scaleY,
	};
}

// Apply a transform to a shape's final (x,y,cx,cy). No-op for identity.
function applyTransform(shape: ShapeLike, t: Transform): void {
	if (t === IDENTITY) return;
	const nx = t.offsetX + shape.x * t.scaleX;
	const ny = t.offsetY + shape.y * t.scaleY;
	shape.x = nx;
	shape.y = ny;
	shape.cx = shape.cx * t.scaleX;
	shape.cy = shape.cy * t.scaleY;
}

function walkSpTree(container: Element, out: ShapeLike[], ctx: SlideParseContext, t: Transform = IDENTITY): void {
	for (const child of Array.from(container.children)) {
		if (child.namespaceURI !== A_NS.p) continue;
		// Per-shape isolation: one mangled shape shouldn't abort parsing of
		// its siblings. Drop the shape + warn; grpSp failures drop the whole
		// group (but not surrounding shapes).
		try {
			if (child.localName === "sp") {
				const s = parseShape(child, ctx);
				if (s) { applyTransform(s, t); out.push(s); }
			} else if (child.localName === "pic") {
				const p = parsePic(child, ctx);
				if (p) { applyTransform(p, t); out.push(p); }
			} else if (child.localName === "grpSp") {
				// Nested group — compose the group's own transform with the
				// accumulated one before recursing into children.
				const grpT = composeTransforms(t, transformFromGrpSp(child));
				walkSpTree(child, out, ctx, grpT);
			} else if (child.localName === "graphicFrame") {
				const gfShape = parseGraphicFrame(child, ctx);
				if (gfShape) { applyTransform(gfShape, t); out.push(gfShape); }
			} else if (child.localName === "cxnSp") {
				// Connector shapes — lines, arrows, etc. We reuse parseShape but
				// cxnSp has no <a:txBody> and typically renders as a line.
				const s = parseShape(child, ctx);
				if (s) { applyTransform(s, t); out.push(s); }
			}
		} catch (err) {
			if (typeof console !== "undefined") console.warn(`[pptxjs] shape <${child.localName}> failed to parse, dropping:`, err);
		}
	}
}

function parseGraphicFrame(gf: Element, ctx: SlideParseContext): ShapeLike | null {
	// <p:xfrm> lives directly under <p:graphicFrame> (not <p:spPr>).
	const xfrm = firstChildNS(gf, A_NS.p, "xfrm");
	const off = xfrm && firstChildNS(xfrm, A_NS.a, "off");
	const ext = xfrm && firstChildNS(xfrm, A_NS.a, "ext");
	const x = off ? Number(off.getAttribute("x")) || 0 : 0;
	const y = off ? Number(off.getAttribute("y")) || 0 : 0;
	const cx = ext ? Number(ext.getAttribute("cx")) || 0 : 0;
	const cy = ext ? Number(ext.getAttribute("cy")) || 0 : 0;
	const rotation60k = xfrm ? Number(xfrm.getAttribute("rot")) || 0 : 0;
	const flipH = xfrm?.getAttribute("flipH") === "1";
	const flipV = xfrm?.getAttribute("flipV") === "1";

	const nvGraphicFramePr = firstChildNS(gf, A_NS.p, "nvGraphicFramePr");
	const nv = parseNonVisualProps(nvGraphicFramePr);

	const graphic = firstChildNS(gf, A_NS.a, "graphic");
	const graphicData = graphic && firstChildNS(graphic, A_NS.a, "graphicData");
	if (!graphicData) return null;
	const uri = graphicData.getAttribute("uri") ?? "";

	if (uri === URI_TABLE) {
		const tbl = firstChildNS(graphicData, A_NS.a, "tbl");
		if (!tbl) return null;

		const colWidthsEmu: number[] = [];
		const tblGrid = firstChildNS(tbl, A_NS.a, "tblGrid");
		if (tblGrid) {
			for (const col of childrenNS(tblGrid, A_NS.a, "gridCol")) {
				colWidthsEmu.push(Number(col.getAttribute("w")) || 0);
			}
		}

		// Table-level properties: boolean flags that toggle band application,
		// and the styleId UUID pointing into ppt/tableStyles.xml.
		const tblPr = firstChildNS(tbl, A_NS.a, "tblPr");
		const tableFlags = emptyTableFlags();
		let styleId: string | null = null;
		if (tblPr) {
			tableFlags.firstRow = tblPr.getAttribute("firstRow") === "1";
			tableFlags.firstCol = tblPr.getAttribute("firstCol") === "1";
			tableFlags.lastRow = tblPr.getAttribute("lastRow") === "1";
			tableFlags.lastCol = tblPr.getAttribute("lastCol") === "1";
			tableFlags.bandRow = tblPr.getAttribute("bandRow") === "1";
			tableFlags.bandCol = tblPr.getAttribute("bandCol") === "1";
			const styleIdEl = firstChildNS(tblPr, A_NS.a, "tableStyleId");
			if (styleIdEl) {
				const txt = styleIdEl.textContent?.trim();
				if (txt) styleId = txt;
			}
		}

		const rows: TableRow[] = [];
		for (const trEl of childrenNS(tbl, A_NS.a, "tr")) {
			rows.push(parseTableRow(trEl, ctx));
		}

		return {
			kind: 'table',
			x, y, cx, cy,
			colWidthsEmu,
			rows,
			styleId,
			tableFlags,
			name: nv.name,
			title: nv.title,
			alt: nv.descr,
			rotation60k,
			flipH,
			flipV,
		};
	}

	if (uri === URI_CHART) {
		let chartRId: string | null = null;
		for (const c of Array.from(graphicData.children)) {
			if (c.localName === "chart") {
				chartRId = c.getAttributeNS(A_NS.r, "id");
				break;
			}
		}
		const out: ChartFallbackShape = {
			kind: 'chart-fallback',
			x, y, cx, cy,
			chartRId,
			src: null,
			name: nv.name,
			title: nv.title,
			alt: nv.descr,
			rotation60k,
			flipH,
			flipV,
		};
		return out;
	}

	if (uri === URI_DIAGRAM) {
		let dataRId: string | null = null;
		for (const c of Array.from(graphicData.children)) {
			if (c.localName === "relIds") {
				dataRId = c.getAttributeNS(A_NS.r, "dm");
				break;
			}
		}
		const out: SmartArtFallbackShape = {
			kind: 'smartart-fallback',
			x, y, cx, cy,
			dataRId,
			name: nv.name,
			title: nv.title,
			alt: nv.descr,
			rotation60k,
			flipH,
			flipV,
		};
		return out;
	}

	// Unknown graphicData — silent skip.
	// eslint-disable-next-line no-console
	console.debug(`[pptxjs] unknown graphicFrame uri: ${uri}`);
	return null;
}

export function parseShapesFromSpTree(spTree: Element, ctx: SlideParseContext): ShapeLike[] {
	const out: ShapeLike[] = [];
	walkSpTreeStatic(spTree, out, ctx);
	return out;
}

function parseTableRow(tr: Element, ctx: SlideParseContext): TableRow {
	const heightEmu = Number(tr.getAttribute("h")) || 0;
	const cells: TableCell[] = [];
	for (const tcEl of childrenNS(tr, A_NS.a, "tc")) {
		cells.push(parseTableCell(tcEl, ctx));
	}
	return { heightEmu, cells };
}

function parseTableCell(tc: Element, ctx: SlideParseContext): TableCell {
	const gridSpan = Number(tc.getAttribute("gridSpan")) || 1;
	const rowSpan = Number(tc.getAttribute("rowSpan")) || 1;
	const hMerge = tc.getAttribute("hMerge") === "1";
	const vMerge = tc.getAttribute("vMerge") === "1";

	const tcPr = firstChildNS(tc, A_NS.a, "tcPr");

	// Fill: honour any <a:fill>-kind child of <a:tcPr>, not just solidFill.
	let fill: Fill | null = null;
	if (tcPr) {
		for (const child of Array.from(tcPr.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			const ln = child.localName;
			if (
				ln === 'solidFill' || ln === 'noFill' || ln === 'gradFill'
				|| ln === 'blipFill' || ln === 'pattFill'
			) {
				fill = parseFillElement(child, ctx.clrMap, ctx.theme);
				break;
			}
		}
	}

	// Cell insets (marL/marR/marT/marB on <a:tcPr>) and anchor.
	let insetsEmu: { l: number; r: number; t: number; b: number } | null = null;
	let anchor: 't' | 'ctr' | 'b' | null = null;
	if (tcPr) {
		const marL = tcPr.getAttribute("marL");
		const marR = tcPr.getAttribute("marR");
		const marT = tcPr.getAttribute("marT");
		const marB = tcPr.getAttribute("marB");
		if (marL != null || marR != null || marT != null || marB != null) {
			insetsEmu = {
				l: marL != null ? Number(marL) || 0 : 91440,
				r: marR != null ? Number(marR) || 0 : 91440,
				t: marT != null ? Number(marT) || 0 : 45720,
				b: marB != null ? Number(marB) || 0 : 45720,
			};
		}
		const anch = tcPr.getAttribute("anchor");
		if (anch === 't' || anch === 'ctr' || anch === 'b') anchor = anch;
	}

	// Per-side borders: <a:lnL>, <a:lnR>, <a:lnT>, <a:lnB>, plus
	// <a:lnTlToBr>, <a:lnBlToTr> for diagonals. Each is a full <a:ln>.
	const borders: TableCellBorders = {
		l: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnL") : null, ctx.clrMap, ctx.theme),
		r: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnR") : null, ctx.clrMap, ctx.theme),
		t: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnT") : null, ctx.clrMap, ctx.theme),
		b: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnB") : null, ctx.clrMap, ctx.theme),
		tlbr: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnTlToBr") : null, ctx.clrMap, ctx.theme),
		blTr: parseLine(tcPr ? firstChildNS(tcPr, A_NS.a, "lnBlToTr") : null, ctx.clrMap, ctx.theme),
	};

	const txBody = firstChildNS(tc, A_NS.a, "txBody");
	const paragraphs: Paragraph[] = [];
	if (txBody) {
		const cellLstStyle = firstChildNS(txBody, A_NS.a, "lstStyle");
		const cellLevelStyles = parseLevelStyles(cellLstStyle, ctx.clrMap, ctx.theme);
		for (const pEl of childrenNS(txBody, A_NS.a, "p")) {
			// Table cells aren't placeholders, so no ph-based inheritance.
			paragraphs.push(parseParagraph(pEl, null, null, cellLevelStyles, ctx));
		}
	}

	return { paragraphs, fill, gridSpan, rowSpan, hMerge, vMerge, insetsEmu, anchor, borders };
}

function parsePic(pic: Element, ctx: SlideParseContext): PicShape | null {
	const spPr = firstChildNS(pic, A_NS.p, "spPr");
	const frame = readXfrm(spPr);
	const blipFill = firstChildNS(pic, A_NS.p, "blipFill");
	const blip = blipFill && firstChildNS(blipFill, A_NS.a, "blip");
	const rId = blip?.getAttributeNS(A_NS.r, "embed") ?? null;
	const src = rId ? (ctx.embedUrls.get(rId) ?? null) : null;

	// <a:srcRect l="" t="" r="" b=""/>. Missing attrs default to 0.
	let srcRectPermille: SrcRect | null = null;
	const srcRectEl = blipFill ? firstChildNS(blipFill, A_NS.a, "srcRect") : null;
	if (srcRectEl) {
		const l = Number(srcRectEl.getAttribute("l")) || 0;
		const t = Number(srcRectEl.getAttribute("t")) || 0;
		const r = Number(srcRectEl.getAttribute("r")) || 0;
		const b = Number(srcRectEl.getAttribute("b")) || 0;
		if (l !== 0 || t !== 0 || r !== 0 || b !== 0) {
			srcRectPermille = { l, t, r, b };
		}
	}

	// <a:stretch> vs <a:tile>. If neither, OOXML defaults to stretch.
	const tileEl = blipFill ? firstChildNS(blipFill, A_NS.a, "tile") : null;
	const stretchEl = blipFill ? firstChildNS(blipFill, A_NS.a, "stretch") : null;
	const stretch = tileEl ? false : (stretchEl ? true : true);
	let tile: TileInfo | null = null;
	if (tileEl) {
		tile = {
			sxPermille: Number(tileEl.getAttribute("sx")) || 100000,
			syPermille: Number(tileEl.getAttribute("sy")) || 100000,
			txPermille: Number(tileEl.getAttribute("tx")) || 0,
			tyPermille: Number(tileEl.getAttribute("ty")) || 0,
		};
	}

	// Blip effect children.
	let alphaPermille: number | null = null;
	let lumBrightPermille: number | null = null;
	let lumContrastPermille: number | null = null;
	let grayscale = false;
	let biLevelPermille: number | null = null;
	let duotone: [string, string] | null = null;
	if (blip) {
		for (const child of Array.from(blip.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			switch (child.localName) {
				case 'alphaModFix': {
					const amt = child.getAttribute("amt");
					// Per ECMA-376 the default when `amt` is omitted is 100000
					// (fully opaque) — effectively a no-op, so only record when
					// the element is materially constraining.
					alphaPermille = amt !== null ? Number(amt) : 100000;
					break;
				}
				case 'lum': {
					const bright = child.getAttribute("bright");
					const contrast = child.getAttribute("contrast");
					if (bright !== null) lumBrightPermille = Number(bright);
					if (contrast !== null) lumContrastPermille = Number(contrast);
					// Presence alone (no attrs) is a no-op — skip.
					break;
				}
				case 'grayscl':
					grayscale = true;
					break;
				case 'biLevel': {
					const thresh = child.getAttribute("thresh");
					biLevelPermille = thresh !== null ? Number(thresh) : 50000;
					break;
				}
				case 'duotone': {
					const colorEls = Array.from(child.children).filter(
						c => c.namespaceURI === A_NS.a,
					);
					const resolved = colorEls
						.map(c => resolveColorElement(c, ctx.clrMap, ctx.theme))
						.filter((r): r is NonNullable<typeof r> => !!r && !!r.colorHex);
					if (resolved.length >= 2) {
						duotone = [resolved[0].colorHex, resolved[1].colorHex];
					}
					break;
				}
			}
		}
	}

	const nvPicPr = firstChildNS(pic, A_NS.p, "nvPicPr");
	const nv = parseNonVisualProps(nvPicPr);
	const alt = nv.descr ?? nv.name ?? "";
	const hyperlinkRId = hyperlinkFromCNvPr(nvPicPr);

	if (!frame && !src) return null;
	return {
		kind: 'pic',
		x: frame?.x ?? 0,
		y: frame?.y ?? 0,
		cx: frame?.cx ?? 0,
		cy: frame?.cy ?? 0,
		src,
		alt,
		name: nv.name,
		title: nv.title,
		hyperlinkRId,
		rotation60k: frame?.rotation60k ?? 0,
		flipH: frame?.flipH ?? false,
		flipV: frame?.flipV ?? false,
		srcRectPermille,
		stretch,
		tile,
		alphaPermille,
		lumBrightPermille,
		lumContrastPermille,
		grayscale,
		biLevelPermille,
		duotone,
	};
}

function parseShape(sp: Element, ctx: SlideParseContext): Shape | null {
	const spPr = firstChildNS(sp, A_NS.p, "spPr");
	let frame = readXfrm(spPr);

	// Resolve the placeholder type (if any) — drives both geometry and
	// default text-style inheritance.
	const nvSpPr = firstChildNS(sp, A_NS.p, "nvSpPr")
		?? firstChildNS(sp, A_NS.p, "nvCxnSpPr");
	const nvPr = nvSpPr && firstChildNS(nvSpPr, A_NS.p, "nvPr");
	const ph = nvPr && firstChildNS(nvPr, A_NS.p, "ph");
	const phType = ph ? normalizePhType(ph.getAttribute("type")) : null;
	const phIdx = ph?.getAttribute("idx") ?? "0";
	const phKey = phType ? `${phType}:${phIdx}` : null;

	if (!frame && phKey) {
		const layoutInfo = ctx.layout.byKey.get(phKey) ?? ctx.layout.byType.get(phType!);
		const masterInfo = ctx.master.byKey.get(phKey) ?? ctx.master.byType.get(phType!);
		frame = layoutInfo?.frame ?? masterInfo?.frame ?? null;
	}

	const txBody = firstChildNS(sp, A_NS.p, "txBody");
	const shapeLstStyle = txBody ? firstChildNS(txBody, A_NS.a, "lstStyle") : null;
	const shapeLevelStyles = parseLevelStyles(shapeLstStyle, ctx.clrMap, ctx.theme);
	const bodyPr = txBody ? parseBodyPr(firstChildNS(txBody, A_NS.a, "bodyPr")) : null;

	const paragraphs: Paragraph[] = [];
	if (txBody) {
		for (const pEl of childrenNS(txBody, A_NS.a, "p")) {
			paragraphs.push(parseParagraph(pEl, phType, phKey, shapeLevelStyles, ctx));
		}
	}

	// Shape fill & border. Non-placeholder shapes often carry these directly;
	// placeholders can also override the inherited look. We look for any fill
	// child kind, not just solidFill.
	let fill: Fill | null = null;
	if (spPr) {
		for (const child of Array.from(spPr.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			const ln = child.localName;
			if (
				ln === 'solidFill' || ln === 'noFill' || ln === 'gradFill'
				|| ln === 'blipFill' || ln === 'pattFill'
			) {
				fill = parseFillElement(child, ctx.clrMap, ctx.theme);
				break;
			}
		}
	}
	const line = parseLine(spPr ? firstChildNS(spPr, A_NS.a, "ln") : null, ctx.clrMap, ctx.theme);
	const custGeom = spPr ? parseCustGeom(firstChildNS(spPr, A_NS.a, "custGeom")) : null;
	const presetGeom = spPr ? parsePresetGeom(firstChildNS(spPr, A_NS.a, "prstGeom")) : null;
	const effects = parseEffectsFromSpPr(spPr, ctx.clrMap, ctx.theme);

	const nv = parseNonVisualProps(nvSpPr);
	const hyperlinkRId = hyperlinkFromCNvPr(nvSpPr);

	if (!frame && paragraphs.length === 0 && !fill && !line && !custGeom && !presetGeom && !effects) return null;
	return {
		kind: 'shape',
		x: frame?.x ?? 0,
		y: frame?.y ?? 0,
		cx: frame?.cx ?? 0,
		cy: frame?.cy ?? 0,
		fill,
		line,
		paragraphs,
		custGeom,
		presetGeom,
		effects,
		name: nv.name,
		title: nv.title,
		alt: nv.descr,
		hyperlinkRId,
		rotation60k: frame?.rotation60k ?? 0,
		flipH: frame?.flipH ?? false,
		flipV: frame?.flipV ?? false,
		phType,
		bodyPr,
	};
}

function parseCustGeom(custGeom: Element | null): CustGeom | null {
	if (!custGeom) return null;
	const pathLst = firstChildNS(custGeom, A_NS.a, "pathLst");
	if (!pathLst) return null;
	const paths = childrenNS(pathLst, A_NS.a, "path");
	if (paths.length === 0) return null;

	const firstW = Number(paths[0].getAttribute("w")) || 0;
	const firstH = Number(paths[0].getAttribute("h")) || 0;
	if (firstW === 0 && firstH === 0) return null;

	let firstFillMode: 'none' | 'norm' | 'darken' | 'lighten' = 'norm';
	let anyClosed = false;
	let d = "";
	const readPt = (pt: Element) => ({
		x: Number(pt.getAttribute("x")) || 0,
		y: Number(pt.getAttribute("y")) || 0,
	});

	for (let pi = 0; pi < paths.length; pi++) {
		const path = paths[pi];
		const fillAttr = path.getAttribute("fill");
		if (pi === 0) {
			if (fillAttr === 'none' || fillAttr === 'norm' || fillAttr === 'darken' || fillAttr === 'lighten') {
				firstFillMode = fillAttr;
			}
		}
		// Track pen position so arcTo (OOXML's relative-to-current-point
		// elliptical arc) can derive its start point.
		let cur: { x: number; y: number } | null = null;
		for (const cmd of Array.from(path.children)) {
			if (cmd.namespaceURI !== A_NS.a) continue;
			if (cmd.localName === "moveTo") {
				const pt = firstChildNS(cmd, A_NS.a, "pt");
				if (pt) { const p = readPt(pt); d += `M ${p.x} ${p.y} `; cur = p; }
			} else if (cmd.localName === "lnTo") {
				const pt = firstChildNS(cmd, A_NS.a, "pt");
				if (pt) { const p = readPt(pt); d += `L ${p.x} ${p.y} `; cur = p; }
			} else if (cmd.localName === "cubicBezTo") {
				// Three <a:pt> children: two control points + end.
				const pts = Array.from(cmd.children)
					.filter(c => c.namespaceURI === A_NS.a && c.localName === "pt")
					.map(readPt);
				if (pts.length >= 3) {
					d += `C ${pts[0].x} ${pts[0].y} ${pts[1].x} ${pts[1].y} ${pts[2].x} ${pts[2].y} `;
					cur = pts[2];
				}
			} else if (cmd.localName === "quadBezTo") {
				const pts = Array.from(cmd.children)
					.filter(c => c.namespaceURI === A_NS.a && c.localName === "pt")
					.map(readPt);
				if (pts.length >= 2) {
					d += `Q ${pts[0].x} ${pts[0].y} ${pts[1].x} ${pts[1].y} `;
					cur = pts[1];
				}
			} else if (cmd.localName === "arcTo") {
				const seg = arcToSvg(cmd, cur);
				if (seg) {
					d += seg.d;
					cur = seg.end;
				}
			} else if (cmd.localName === "close") {
				d += "Z ";
				anyClosed = true;
			}
		}
	}
	if (!d) return null;
	return { pathW: firstW, pathH: firstH, d: d.trim(), closed: anyClosed, fillMode: firstFillMode };
}

// Convert <a:arcTo wR hR stAng swAng/> to an SVG elliptical arc command.
//
// OOXML arcTo semantics (ECMA-376 §20.1.9.3):
//   - `wR` / `hR` are the ellipse radii in path-local units.
//   - `stAng` / `swAng` are angles in 60000ths of a degree; 0° points east
//     (3 o'clock) and positive angles sweep clockwise — the same sign
//     convention SVG uses in screen space, where y grows downward.
//   - The arc's centre is derived from the current pen position:
//     (cx, cy) = (curX - wR*cos(stAng), curY - hR*sin(stAng)).
//   - The endpoint is (cx + wR*cos(stAng+swAng), cy + hR*sin(stAng+swAng)).
//
// SVG's `A rx ry x-axis-rot large-arc-flag sweep-flag x y` then needs
// `large-arc-flag = |swAng| > 180°` and `sweep-flag = swAng > 0` (CW).
// When there's no current pen position we assume (0, 0) — the OOXML
// schema in practice always precedes arcTo with a moveTo.
function arcToSvg(cmd: Element, cur: { x: number; y: number } | null): { d: string; end: { x: number; y: number } } | null {
	const wR = Number(cmd.getAttribute("wR")) || 0;
	const hR = Number(cmd.getAttribute("hR")) || 0;
	const stAng60k = Number(cmd.getAttribute("stAng")) || 0;
	const swAng60k = Number(cmd.getAttribute("swAng")) || 0;
	if (wR === 0 || hR === 0) return null;
	const toRad = (a60k: number) => (a60k / 60000) * Math.PI / 180;
	const stA = toRad(stAng60k);
	const enA = toRad(stAng60k + swAng60k);
	const curX = cur?.x ?? 0;
	const curY = cur?.y ?? 0;
	const cx = curX - wR * Math.cos(stA);
	const cy = curY - hR * Math.sin(stA);
	const endX = cx + wR * Math.cos(enA);
	const endY = cy + hR * Math.sin(enA);
	const absSwDeg = Math.abs(swAng60k / 60000);
	const largeArc = absSwDeg > 180 ? 1 : 0;
	const sweep = swAng60k > 0 ? 1 : 0;
	return {
		d: `A ${wR} ${hR} 0 ${largeArc} ${sweep} ${endX} ${endY} `,
		end: { x: endX, y: endY },
	};
}

// Parse <a:prstGeom prst="…"> with its optional <a:avLst> of adjust values.
// Each <a:gd name="adjN" fmla="val 12345"/> contributes adjN → 12345 (EMU-free,
// per-100000 units). fmla values not of the simple "val N" form are ignored.
function parsePresetGeom(prstGeom: Element | null): { name: string; avLst: Map<string, number> } | null {
	if (!prstGeom) return null;
	const name = prstGeom.getAttribute("prst");
	if (!name) return null;
	const avLst = new Map<string, number>();
	const avLstEl = firstChildNS(prstGeom, A_NS.a, "avLst");
	if (avLstEl) {
		for (const gd of childrenNS(avLstEl, A_NS.a, "gd")) {
			const gdName = gd.getAttribute("name");
			const fmla = gd.getAttribute("fmla");
			if (!gdName || !fmla) continue;
			// We only handle "val N" formulas (the common case). Anything else
			// — "*/ a b c", "pin 0 X Y" etc. — is left unset so presetToSvgPath
			// falls back to its default.
			const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(fmla.trim());
			if (m) avLst.set(gdName, Number(m[1]));
		}
	}
	return { name, avLst };
}

function readXfrm(spPr: Element | null): PlaceholderFrame | null {
	if (!spPr) return null;
	const xfrm = firstChildNS(spPr, A_NS.a, "xfrm");
	if (!xfrm) return null;
	const off = firstChildNS(xfrm, A_NS.a, "off");
	const ext = firstChildNS(xfrm, A_NS.a, "ext");
	if (!off || !ext) return null;
	const cx = Number(ext.getAttribute("cx")) || 0;
	const cy = Number(ext.getAttribute("cy")) || 0;
	if (cx === 0 && cy === 0) return null;
	const rotation60k = Number(xfrm.getAttribute("rot")) || 0;
	const flipH = xfrm.getAttribute("flipH") === "1";
	const flipV = xfrm.getAttribute("flipV") === "1";
	return {
		x: Number(off.getAttribute("x")) || 0,
		y: Number(off.getAttribute("y")) || 0,
		cx,
		cy,
		rotation60k,
		flipH,
		flipV,
	};
}

function parseParagraph(
	pEl: Element,
	phType: string | null,
	phKey: string | null,
	shapeLevelStyles: LevelStyles,
	ctx: SlideParseContext,
): Paragraph {
	const pPr = firstChildNS(pEl, A_NS.a, "pPr");
	const level = pPr?.getAttribute("lvl") ? Number(pPr.getAttribute("lvl")) : 0;

	// Build base styles — both paragraph- and run-level — by stacking
	// inheritance sources from weakest to strongest:
	//   master txStyles → master placeholder → layout placeholder
	//   → shape lstStyle → paragraph pPr (+ its defRPr for runs).
	let baseRun = emptyRunStyle();
	let basePara = emptyParaStyle();

	const applyLevel = (lvl: { run: RunStyle | null; para: ParaStyle | null } | null) => {
		if (!lvl) return;
		if (lvl.run) baseRun = mergeRunStyle(baseRun, lvl.run);
		if (lvl.para) basePara = mergeParaStyle(basePara, lvl.para);
	};

	if (phType) {
		const bucket = phTypeToStyleBucket(phType);
		applyLevel(ctx.masterTextStyles[bucket][level]);
	}
	if (phKey && phType) {
		const masterInfo = ctx.master.byKey.get(phKey) ?? ctx.master.byType.get(phType);
		applyLevel(masterInfo?.levelStyles[level] ?? null);
		const layoutInfo = ctx.layout.byKey.get(phKey) ?? ctx.layout.byType.get(phType);
		applyLevel(layoutInfo?.levelStyles[level] ?? null);
	}
	applyLevel(shapeLevelStyles[level] ?? null);

	// Paragraph's own pPr overrides (attributes + children like <a:buChar>).
	if (pPr) basePara = mergeParaStyle(basePara, parseParaProps(pPr, ctx.clrMap, ctx.theme));
	const pDefRPr = pPr ? firstChildNS(pPr, A_NS.a, "defRPr") : null;
	if (pDefRPr) baseRun = mergeRunStyle(baseRun, parseRunProps(pDefRPr, ctx.clrMap, ctx.theme));

	const runs: Run[] = [];
	for (const rEl of Array.from(pEl.children)) {
		if (rEl.namespaceURI !== A_NS.a) continue;
		const ln = rEl.localName;
		if (ln === 'r') {
			runs.push(buildTextRun(rEl, baseRun, ctx));
		} else if (ln === 'br') {
			runs.push({ kind: 'break' });
		} else if (ln === 'fld') {
			runs.push(buildFieldRun(rEl, baseRun, ctx));
		}
	}
	return { level, style: basePara, runs };
}

function resolveRunStyle(style: RunStyle, ctx: SlideParseContext): RunStyle {
	if (!style.colorHex && style.colorSchemeSlot) {
		const base = resolveSchemeClr(style.colorSchemeSlot, ctx.clrMap, ctx.theme);
		if (base) {
			style.colorHex = style.colorMods ? applyMods(base, style.colorMods) : base;
		}
	}
	style.fontFamily = resolveTypeface(style.fontFamily, ctx.theme);
	return style;
}

function buildTextRun(rEl: Element, baseRun: RunStyle, ctx: SlideParseContext): TextRun {
	const tEl = firstChildNS(rEl, A_NS.a, "t");
	const rPr = firstChildNS(rEl, A_NS.a, "rPr");
	const runOverrides = parseRunProps(rPr, ctx.clrMap, ctx.theme);
	const style = resolveRunStyle(mergeRunStyle(baseRun, runOverrides), ctx);
	return { kind: 'text', text: tEl?.textContent ?? "", style };
}

function buildFieldRun(fldEl: Element, baseRun: RunStyle, ctx: SlideParseContext): FieldRun {
	const fieldType = fldEl.getAttribute("type") ?? "";
	const tEl = firstChildNS(fldEl, A_NS.a, "t");
	const rPr = firstChildNS(fldEl, A_NS.a, "rPr");
	const runOverrides = parseRunProps(rPr, ctx.clrMap, ctx.theme);
	const style = resolveRunStyle(mergeRunStyle(baseRun, runOverrides), ctx);
	return {
		kind: 'field',
		fieldType,
		fallbackText: tEl?.textContent ?? "",
		style,
	};
}
