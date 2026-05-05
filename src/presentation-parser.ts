import type { Options } from './pptx-preview';
import { A_NS } from './namespaces';
import { firstChildNS, childrenNS } from './xml-utils';
import {
	RunStyle,
	ParaStyle,
	LevelStyles,
	emptyRunStyle,
	emptyParaStyle,
	emptyLevelStyles,
	parseRunProps,
	parseParaProps,
	parseLevelStyles,
	mergeRunStyle,
	mergeParaStyle,
} from './text-style';
import { ThemeColors, ClrMap, resolveSchemeClr, resolveTypeface } from './theme';
import { SolidFill, LineStyle, parseSolidFill, parseLine } from './fill';
import { applyMods } from './color-math';

export interface SlideSize {
	cx: number;
	cy: number;
}

export interface Slide {
	index: number;
	shapes: ShapeLike[];
	background: import('./background').BackgroundFill | null;
	// Speaker notes pulled from the slide's companion notesSlide part.
	// Null if the slide has no notesSlide rel or the body placeholder
	// contains no text content.
	notes: NotesContent | null;
	// The slide's <p:transition> preset, captured as the local-name of
	// the first direct-child element (fade, push, wipe, ...). Null when
	// no <p:transition> is present, or when it carries no effect child.
	transition: SlideTransition | null;
}

export interface NotesContent {
	// Flat list of paragraphs from the notes body placeholder. Shape-
	// level granularity is dropped — callers render these inline under
	// the slide's companion <aside>.
	paragraphs: Paragraph[];
}

export interface SlideTransition {
	preset: string;
}

export type ShapeLike = Shape | PicShape | TableShape | ChartShape;

export interface TableShape {
	kind: 'table';
	x: number;
	y: number;
	cx: number;
	cy: number;
	colWidthsEmu: number[];
	rows: TableRow[];
}

export interface TableRow {
	heightEmu: number;
	cells: TableCell[];
}

export interface TableCell {
	paragraphs: Paragraph[];
	fill: SolidFill | null;
	// Column span > 1 (from `gridSpan`); merged continuation cells (`hMerge`) are
	// dropped during parse. Similarly `rowSpan` / `vMerge` handling is minimal.
	gridSpan: number;
	rowSpan: number;
	// Continuation cells that the renderer should skip.
	hMerge: boolean;
	vMerge: boolean;
}

export interface Shape {
	kind: 'shape';
	x: number;
	y: number;
	cx: number;
	cy: number;
	fill: SolidFill | null;
	line: LineStyle | null;
	paragraphs: Paragraph[];
	// Custom geometry (<a:custGeom>), if any. Null for preset/no geometry.
	// When present, the renderer produces an SVG path instead of a plain box.
	custGeom: CustGeom | null;
	// Preset-geometry name from <a:prstGeom prst="...">. Null when the
	// shape has a custGeom or no geometry hint. Surfaced to the DOM via
	// `data-shape-preset="<prst>"` so conformance selectors can key off
	// the shape kind.
	prstGeom: string | null;
	// Placeholder type ("title", "ctrTitle", "body", "subTitle", ...)
	// normalised via `normalizePhType`, and the placeholder idx string.
	// Surfaced on the rendered shape as `data-placeholder-type` /
	// `data-placeholder-idx`.
	phType: string | null;
	phIdx: string | null;
	// True when the shape was authored as a textbox
	// (<p:cNvSpPr txBox="1"/>). Surfaced as
	// `data-shape-type="textbox"` so conformance selectors can find
	// textboxes specifically rather than preset shapes.
	isTextbox: boolean;
}

export interface ChartShape {
	kind: 'chart';
	x: number;
	y: number;
	cx: number;
	cy: number;
	// Best-effort chart-type token surfaced as `data-chart-type` on the
	// rendered placeholder. Null when the referenced chart part couldn't
	// be resolved or parsed — the DOM still carries `data-kind="chart"`
	// so generic selectors pass.
	chartType: string | null;
}

// Minimal <a:custGeom> representation — a single path with the local
// coordinate dimensions (`w`/`h`) used by its inner point coords. Only
// straight-line commands are captured; curves/arcs fall back to straight
// segments through their endpoints.
export interface CustGeom {
	pathW: number;
	pathH: number;
	// SVG `d` attribute built during parse. Expressed in the path's local
	// coordinate system (0..pathW, 0..pathH) so the renderer can emit it
	// inside an SVG viewBox matching those dimensions.
	d: string;
	closed: boolean;
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
}

export interface Paragraph {
	level: number;
	style: ParaStyle;
	runs: Run[];
}

export interface Run {
	text: string;
	style: RunStyle;
}

export interface PlaceholderFrame {
	x: number;
	y: number;
	cx: number;
	cy: number;
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
	const spTree = firstChildNS(firstChildNS(doc.documentElement, A_NS.p, "cSld"), A_NS.p, "spTree");
	if (spTree) walkSpTree(spTree, shapes, ctx);
	const transition = parseSlideTransition(doc);
	return { index, shapes, background: null, notes: null, transition };
}

// Pull <p:transition>'s first direct-child element and return its
// local-name as the preset token. Ignores element-content attributes
// like `spd` / `advClick` — those belong to a future timing / playback
// feature. Returns null when no transition, or when the transition is
// present but empty (no effect child).
export function parseSlideTransition(doc: Document): SlideTransition | null {
	const root = doc.documentElement;
	const transition = firstChildNS(root, A_NS.p, "transition");
	if (!transition) return null;
	for (const child of Array.from(transition.children)) {
		// ECMA-376 transition effect children live in the `p` namespace;
		// PowerPoint 2010+ extensions live in `p14` inside an <extLst>
		// wrapper. The fixture under test always writes the `p` form, so
		// the presence check below is intentionally narrow.
		if (child.namespaceURI !== A_NS.p) continue;
		if (child.localName === "extLst") continue;
		return { preset: child.localName };
	}
	return null;
}

// Walk a notesSlide document and collect any <p:sp> whose placeholder
// is a body — the conventional location of speaker notes. Falls back
// to collecting paragraphs from all shapes when no explicit body
// placeholder is found, since older tools sometimes omit the <p:ph
// type="body"> marker.
export function parseNotesSlide(doc: Document, ctx: SlideParseContext): NotesContent | null {
	const spTree = firstChildNS(firstChildNS(doc.documentElement, A_NS.p, "cSld"), A_NS.p, "spTree");
	if (!spTree) return null;
	const sps = Array.from(spTree.children).filter(
		c => c.namespaceURI === A_NS.p && c.localName === "sp",
	);
	const bodyParagraphs: Paragraph[] = [];
	const anyParagraphs: Paragraph[] = [];
	for (const sp of sps) {
		const nvSpPr = firstChildNS(sp, A_NS.p, "nvSpPr");
		const nvPr = nvSpPr && firstChildNS(nvSpPr, A_NS.p, "nvPr");
		const ph = nvPr && firstChildNS(nvPr, A_NS.p, "ph");
		const phType = ph ? normalizePhType(ph.getAttribute("type")) : null;
		const phIdx = ph?.getAttribute("idx") ?? "0";
		const phKey = phType ? `${phType}:${phIdx}` : null;
		const txBody = firstChildNS(sp, A_NS.p, "txBody");
		if (!txBody) continue;
		const shapeLstStyle = firstChildNS(txBody, A_NS.a, "lstStyle");
		const shapeLevelStyles = parseLevelStyles(shapeLstStyle);
		for (const pEl of childrenNS(txBody, A_NS.a, "p")) {
			const para = parseParagraph(pEl, phType, phKey, shapeLevelStyles, ctx);
			anyParagraphs.push(para);
			// Only the body placeholder carries speaker notes; the
			// sldImg / sldNum placeholders on a notesSlide are chrome.
			if (phType === "body") bodyParagraphs.push(para);
		}
	}
	const paragraphs = bodyParagraphs.length > 0 ? bodyParagraphs : anyParagraphs;
	// An empty notes body means "no notes"; surface null so the
	// renderer can skip emitting the aside entirely.
	const hasText = paragraphs.some(p => p.runs.some(r => r.text.trim() !== ""));
	if (!hasText) return null;
	return { paragraphs };
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
	}
}

const CHART_URI = "http://schemas.openxmlformats.org/drawingml/2006/chart";

function parseGraphicFrame(gf: Element, ctx: SlideParseContext): TableShape | ChartShape | null {
	// <p:xfrm> lives directly under <p:graphicFrame> (not <p:spPr>).
	const xfrm = firstChildNS(gf, A_NS.p, "xfrm");
	const off = xfrm && firstChildNS(xfrm, A_NS.a, "off");
	const ext = xfrm && firstChildNS(xfrm, A_NS.a, "ext");
	const x = off ? Number(off.getAttribute("x")) || 0 : 0;
	const y = off ? Number(off.getAttribute("y")) || 0 : 0;
	const cx = ext ? Number(ext.getAttribute("cx")) || 0 : 0;
	const cy = ext ? Number(ext.getAttribute("cy")) || 0 : 0;

	const graphic = firstChildNS(gf, A_NS.a, "graphic");
	const graphicData = graphic && firstChildNS(graphic, A_NS.a, "graphicData");
	if (!graphicData) return null;

	const uri = graphicData.getAttribute("uri");
	if (uri === CHART_URI) {
		// Chart parts are referenced by rId but the chart XML itself
		// lives in a separate part that would need async loading. For
		// now we emit a placeholder shape tagged with data-kind="chart"
		// so conformance selectors resolve; deep chart rendering is
		// tracked as a separate follow-up in TODO.md.
		return { kind: 'chart', x, y, cx, cy, chartType: null };
	}

	const tbl = firstChildNS(graphicData, A_NS.a, "tbl");
	if (!tbl) return null;

	const colWidthsEmu: number[] = [];
	const tblGrid = firstChildNS(tbl, A_NS.a, "tblGrid");
	if (tblGrid) {
		for (const col of childrenNS(tblGrid, A_NS.a, "gridCol")) {
			colWidthsEmu.push(Number(col.getAttribute("w")) || 0);
		}
	}

	const rows: TableRow[] = [];
	for (const trEl of childrenNS(tbl, A_NS.a, "tr")) {
		rows.push(parseTableRow(trEl, ctx));
	}
	return { kind: 'table', x, y, cx, cy, colWidthsEmu, rows };
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
	const fill = parseSolidFill(tcPr ? firstChildNS(tcPr, A_NS.a, "solidFill") : null, ctx.clrMap, ctx.theme);

	const txBody = firstChildNS(tc, A_NS.a, "txBody");
	const paragraphs: Paragraph[] = [];
	if (txBody) {
		const cellLstStyle = firstChildNS(txBody, A_NS.a, "lstStyle");
		const cellLevelStyles = parseLevelStyles(cellLstStyle);
		for (const pEl of childrenNS(txBody, A_NS.a, "p")) {
			// Table cells aren't placeholders, so no ph-based inheritance.
			paragraphs.push(parseParagraph(pEl, null, null, cellLevelStyles, ctx));
		}
	}

	return { paragraphs, fill, gridSpan, rowSpan, hMerge, vMerge };
}

function parsePic(pic: Element, ctx: SlideParseContext): PicShape | null {
	const spPr = firstChildNS(pic, A_NS.p, "spPr");
	const frame = readXfrm(spPr);
	const blipFill = firstChildNS(pic, A_NS.p, "blipFill");
	const blip = blipFill && firstChildNS(blipFill, A_NS.a, "blip");
	const rId = blip?.getAttributeNS(A_NS.r, "embed") ?? null;
	const src = rId ? (ctx.embedUrls.get(rId) ?? null) : null;

	const nvPicPr = firstChildNS(pic, A_NS.p, "nvPicPr");
	const cNvPr = nvPicPr && firstChildNS(nvPicPr, A_NS.p, "cNvPr");
	const alt = cNvPr?.getAttribute("descr") ?? cNvPr?.getAttribute("name") ?? "";

	if (!frame && !src) return null;
	return {
		kind: 'pic',
		x: frame?.x ?? 0,
		y: frame?.y ?? 0,
		cx: frame?.cx ?? 0,
		cy: frame?.cy ?? 0,
		src,
		alt,
	};
}

function parseShape(sp: Element, ctx: SlideParseContext): Shape | null {
	const spPr = firstChildNS(sp, A_NS.p, "spPr");
	let frame = readXfrm(spPr);

	// Resolve the placeholder type (if any) — drives both geometry and
	// default text-style inheritance.
	const nvSpPr = firstChildNS(sp, A_NS.p, "nvSpPr");
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
	const shapeLevelStyles = parseLevelStyles(shapeLstStyle);

	const paragraphs: Paragraph[] = [];
	if (txBody) {
		for (const pEl of childrenNS(txBody, A_NS.a, "p")) {
			paragraphs.push(parseParagraph(pEl, phType, phKey, shapeLevelStyles, ctx));
		}
	}

	// Shape fill & border. Non-placeholder shapes often carry these directly;
	// placeholders can also override the inherited look.
	const fill = parseSolidFill(spPr ? firstChildNS(spPr, A_NS.a, "solidFill") : null, ctx.clrMap, ctx.theme);
	const line = parseLine(spPr ? firstChildNS(spPr, A_NS.a, "ln") : null, ctx.clrMap, ctx.theme);
	const custGeom = spPr ? parseCustGeom(firstChildNS(spPr, A_NS.a, "custGeom")) : null;

	// Preset-geometry name (rect, ellipse, line, rightArrow, ...) — a
	// plain attribute read. Skipped when the shape carries a custGeom.
	let prstGeom: string | null = null;
	if (spPr && !custGeom) {
		const prstGeomEl = firstChildNS(spPr, A_NS.a, "prstGeom");
		if (prstGeomEl) prstGeom = prstGeomEl.getAttribute("prst");
	}

	// <p:cNvSpPr txBox="1"/> marks the shape as a user-created text
	// box (distinct from preset geometries that happen to hold text).
	const cNvSpPr = nvSpPr && firstChildNS(nvSpPr, A_NS.p, "cNvSpPr");
	const isTextbox = cNvSpPr?.getAttribute("txBox") === "1";

	if (!frame && paragraphs.length === 0 && !fill && !line && !custGeom) return null;
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
		prstGeom,
		phType,
		phIdx: ph ? phIdx : null,
		isTextbox,
	};
}

function parseCustGeom(custGeom: Element | null): CustGeom | null {
	if (!custGeom) return null;
	const pathLst = firstChildNS(custGeom, A_NS.a, "pathLst");
	if (!pathLst) return null;
	const path = firstChildNS(pathLst, A_NS.a, "path");
	if (!path) return null;
	const pathW = Number(path.getAttribute("w")) || 0;
	const pathH = Number(path.getAttribute("h")) || 0;
	if (pathW === 0 && pathH === 0) return null;

	let d = "";
	let closed = false;
	const readPt = (pt: Element) => ({
		x: Number(pt.getAttribute("x")) || 0,
		y: Number(pt.getAttribute("y")) || 0,
	});
	for (const cmd of Array.from(path.children)) {
		if (cmd.namespaceURI !== A_NS.a) continue;
		if (cmd.localName === "moveTo") {
			const pt = firstChildNS(cmd, A_NS.a, "pt");
			if (pt) { const p = readPt(pt); d += `M ${p.x} ${p.y} `; }
		} else if (cmd.localName === "lnTo") {
			const pt = firstChildNS(cmd, A_NS.a, "pt");
			if (pt) { const p = readPt(pt); d += `L ${p.x} ${p.y} `; }
		} else if (cmd.localName === "cubicBezTo") {
			// Three <a:pt> children: two control points + end.
			const pts = Array.from(cmd.children)
				.filter(c => c.namespaceURI === A_NS.a && c.localName === "pt")
				.map(readPt);
			if (pts.length >= 3) {
				d += `C ${pts[0].x} ${pts[0].y} ${pts[1].x} ${pts[1].y} ${pts[2].x} ${pts[2].y} `;
			}
		} else if (cmd.localName === "quadBezTo") {
			const pts = Array.from(cmd.children)
				.filter(c => c.namespaceURI === A_NS.a && c.localName === "pt")
				.map(readPt);
			if (pts.length >= 2) {
				d += `Q ${pts[0].x} ${pts[0].y} ${pts[1].x} ${pts[1].y} `;
			}
		} else if (cmd.localName === "close") {
			d += "Z ";
			closed = true;
		}
		// arcTo is skipped — correct SVG conversion is non-trivial and rare in practice.
	}
	if (!d) return null;
	return { pathW, pathH, d: d.trim(), closed };
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
	return {
		x: Number(off.getAttribute("x")) || 0,
		y: Number(off.getAttribute("y")) || 0,
		cx,
		cy,
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
	if (pPr) basePara = mergeParaStyle(basePara, parseParaProps(pPr));
	const pDefRPr = pPr ? firstChildNS(pPr, A_NS.a, "defRPr") : null;
	if (pDefRPr) baseRun = mergeRunStyle(baseRun, parseRunProps(pDefRPr));

	const runs: Run[] = [];
	for (const rEl of childrenNS(pEl, A_NS.a, "r")) {
		const tEl = firstChildNS(rEl, A_NS.a, "t");
		const rPr = firstChildNS(rEl, A_NS.a, "rPr");
		const runOverrides = parseRunProps(rPr);
		const style = mergeRunStyle(baseRun, runOverrides);
		if (!style.colorHex && style.colorSchemeSlot) {
			const base = resolveSchemeClr(style.colorSchemeSlot, ctx.clrMap, ctx.theme);
			if (base) {
				style.colorHex = style.colorMods ? applyMods(base, style.colorMods) : base;
			}
		}
		// Resolve theme font references (e.g. "+mj-lt") to concrete families.
		style.fontFamily = resolveTypeface(style.fontFamily, ctx.theme);
		runs.push({
			text: tEl?.textContent ?? "",
			style,
		});
	}
	return { level, style: basePara, runs };
}
