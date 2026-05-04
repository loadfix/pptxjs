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
import { Fill, LineStyle, parseFillElement, parseLine } from './fill';
import { applyMods } from './color-math';

export interface SlideSize {
	cx: number;
	cy: number;
}

export interface Slide {
	index: number;
	shapes: ShapeLike[];
	background: import('./background').BackgroundFill | null;
}

export type ShapeLike = Shape | PicShape | TableShape;

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
	name: string | null;
	title: string | null;
	hyperlinkRId: string | null;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
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
	const spTree = firstChildNS(firstChildNS(doc.documentElement, A_NS.p, "cSld"), A_NS.p, "spTree");
	if (spTree) walkSpTree(spTree, shapes, ctx);
	return { index, shapes, background: null };
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

function parseGraphicFrame(gf: Element, ctx: SlideParseContext): TableShape | null {
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

	const graphic = firstChildNS(gf, A_NS.a, "graphic");
	const graphicData = graphic && firstChildNS(graphic, A_NS.a, "graphicData");
	const tbl = graphicData && firstChildNS(graphicData, A_NS.a, "tbl");
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

	const nvGraphicFramePr = firstChildNS(gf, A_NS.p, "nvGraphicFramePr");
	const nv = parseNonVisualProps(nvGraphicFramePr);

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

	const nv = parseNonVisualProps(nvSpPr);
	const hyperlinkRId = hyperlinkFromCNvPr(nvSpPr);

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
		name: nv.name,
		title: nv.title,
		alt: nv.descr,
		hyperlinkRId,
		rotation60k: frame?.rotation60k ?? 0,
		flipH: frame?.flipH ?? false,
		flipV: frame?.flipV ?? false,
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
