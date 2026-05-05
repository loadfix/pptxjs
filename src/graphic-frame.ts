// Chart and SmartArt graphic-frame fallback parsing.
//
// PowerPoint stores charts and SmartArt (diagrams) as <p:graphicFrame>
// children of the slide's spTree. We don't implement chart rendering from
// data, nor do we run the SmartArt layout algorithm; instead we surface a
// pre-rendered thumbnail where the package has one (for charts, that's an
// image related to the chart part; for SmartArt, the diagrams/drawingN.xml
// sibling part holds PowerPoint's pre-rendered DrawingML).
//
// Kept separate from presentation-parser so the graphic-frame dispatch in
// that file stays a pure element-to-shape transform — the async package
// lookups needed to resolve targets live here, run from presentation.ts
// after the slide element tree has been walked.

import { OpenXmlPackage, Relationship, resolveRelTarget } from './open-xml-package';
import { imageMimeFromPath } from './mime';
import type { ShapeLike, SlideParseContext } from './presentation-parser';
import { parseChart, type ChartModel } from './chart-parser';

// Relationship types used by chart / SmartArt frames.
export const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
export const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const DIAGRAM_DATA_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData";
// MS-specific but widely used: the pre-rendered DrawingML layout of a
// SmartArt diagram. Set by PowerPoint when it writes the drawing cache.
export const DIAGRAM_DRAWING_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";

// graphicData URI literals.
export const URI_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table";
export const URI_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart";
export const URI_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

// Namespace URIs for the SmartArt drawing cache (dsp) and p-main.
const DSP_NS = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
// DrawingML chart namespace, used to probe chart type from chartN.xml.
const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";

export interface ChartFallbackShape {
	kind: 'chart-fallback';
	x: number;
	y: number;
	cx: number;
	cy: number;
	// rId of the <c:chart r:id="..."/> on the slide — used later (in the
	// presentation loader) to resolve the chart part and pull a cached image.
	chartRId: string | null;
	// Filled in by the post-parse resolver. Null if no cached image found.
	src: string | null;
	// Resolved chart-type label probed from ppt/charts/chartN.xml's first
	// plot-area series element (e.g. "columnClustered", "line", "pie"). Null
	// when the chart part can't be found or the element is unrecognised.
	chartType: string | null;
	// Parsed ChartModel for the chart families we render natively
	// (bar/column/line/pie/doughnut). Null when the chart uses an
	// unsupported family or its part couldn't be read; callers fall back
	// to the "[Chart]" placeholder in that case.
	model: ChartModel | null;
	name: string | null;
	title: string | null;
	alt: string | null;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
}

export interface SmartArtFallbackShape {
	kind: 'smartart-fallback';
	x: number;
	y: number;
	cx: number;
	cy: number;
	// rId of the `r:dm` attr on <dgm:relIds>. Used to locate the diagramData
	// part, whose sibling drawing part holds the pre-rendered shapes.
	dataRId: string | null;
	name: string | null;
	title: string | null;
	alt: string | null;
	rotation60k: number;
	flipH: boolean;
	flipV: boolean;
}

export type GraphicFrameFallbackShape = ChartFallbackShape | SmartArtFallbackShape;

// Resolve a chart's cached image by inspecting its _rels for an image
// relationship. PowerPoint doesn't guarantee one exists — charts are
// usually rendered live from data — but many authoring tools include a
// preview bitmap. Returns a blob URL, or null if nothing usable was found.
export async function findChartFallbackImage(
	pkg: OpenXmlPackage,
	chartPartPath: string,
	mediaUrlCache: Map<string, string>,
): Promise<string | null> {
	const rels = await pkg.loadRelationships(chartPartPath);
	for (const rel of rels.values()) {
		if (rel.type !== IMAGE_REL_TYPE) continue;
		const mediaPath = resolveRelTarget(chartPartPath, rel.target);
		let url = mediaUrlCache.get(mediaPath);
		if (!url) {
			const blob = await pkg.loadBlob(mediaPath, imageMimeFromPath(mediaPath));
			if (blob) {
				url = URL.createObjectURL(blob);
				mediaUrlCache.set(mediaPath, url);
			}
		}
		if (url) return url;
	}
	return null;
}

// Probe the chart part (ppt/charts/chartN.xml) for its first plot-area
// series element and translate it into a chart-type label suitable for
// `data-chart-type="..."`.
//
// The element set is fixed by ECMA-376 — one of barChart, bar3DChart,
// lineChart, line3DChart, pieChart, pie3DChart, doughnutChart, areaChart,
// area3DChart, scatterChart, radarChart, bubbleChart, stockChart,
// surfaceChart, surface3DChart, ofPieChart. For bar/area/line we combine
// with the grouping attribute to emit PowerPoint's conventional labels
// ("columnClustered", "barStacked", "lineStacked100", etc.). Returns null
// when the chart part is missing or holds something we don't recognise.
export async function readChartType(
	pkg: OpenXmlPackage,
	chartPartPath: string,
): Promise<string | null> {
	const doc = await pkg.loadXml(chartPartPath);
	if (!doc) return null;
	// Walk <c:chartSpace> → <c:chart> → <c:plotArea>, then take the first
	// element in the chart-type enumeration.
	const root = doc.documentElement;
	if (!root || root.namespaceURI !== C_NS || root.localName !== "chartSpace") return null;
	const chart = firstChildNS(root, C_NS, "chart");
	if (!chart) return null;
	const plotArea = firstChildNS(chart, C_NS, "plotArea");
	if (!plotArea) return null;
	for (const child of Array.from(plotArea.children)) {
		if (child.namespaceURI !== C_NS) continue;
		const mapped = mapChartTypeElement(child);
		if (mapped) return mapped;
	}
	return null;
}

// Load the chart part and convert it into a ChartModel for inline SVG
// rendering. Returns null when the part is missing, malformed, or uses a
// chart family outside our supported set — the caller should keep the
// "[Chart]" placeholder in that case.
export async function readChartModel(
	pkg: OpenXmlPackage,
	chartPartPath: string,
): Promise<ChartModel | null> {
	const doc = await pkg.loadXml(chartPartPath);
	if (!doc) return null;
	return parseChart(doc);
}

function firstChildNS(parent: Element, ns: string, localName: string): Element | null {
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === ns && c.localName === localName) return c;
	}
	return null;
}

// Map a `<c:*Chart>` element to a chart-type string. Grouping is folded
// into the label for bar/line/area variants following PowerPoint's
// "columnClustered" / "lineStacked100" convention.
function mapChartTypeElement(el: Element): string | null {
	const local = el.localName;
	switch (local) {
		case "barChart":
			return barLabel(el, /* threeD */ false);
		case "bar3DChart":
			return barLabel(el, /* threeD */ true);
		case "lineChart":
			return groupingLabel(el, "line");
		case "line3DChart":
			return groupingLabel(el, "line3D");
		case "pieChart":
			return "pie";
		case "pie3DChart":
			return "pie3D";
		case "doughnutChart":
			return "doughnut";
		case "areaChart":
			return groupingLabel(el, "area");
		case "area3DChart":
			return groupingLabel(el, "area3D");
		case "scatterChart":
			return "scatter";
		case "radarChart":
			return "radar";
		case "bubbleChart":
			return "bubble";
		case "stockChart":
			return "stock";
		case "surfaceChart":
			return "surface";
		case "surface3DChart":
			return "surface3D";
		case "ofPieChart":
			return "ofPie";
		default:
			return null;
	}
}

function barLabel(el: Element, threeD: boolean): string {
	// <c:barDir val="col"|"bar"> decides column vs bar; fall back to "bar"
	// if the attribute is missing (ECMA default is "bar").
	const barDir = firstChildNS(el, C_NS, "barDir");
	const dir = barDir?.getAttribute("val") ?? "bar";
	const base = dir === "col" ? (threeD ? "column3D" : "column") : (threeD ? "bar3D" : "bar");
	return groupingLabel(el, base);
}

function groupingLabel(el: Element, base: string): string {
	const grouping = firstChildNS(el, C_NS, "grouping");
	const val = grouping?.getAttribute("val");
	switch (val) {
		case "clustered":
			return `${base}Clustered`;
		case "stacked":
			return `${base}Stacked`;
		case "percentStacked":
			return `${base}Stacked100`;
		case "standard":
		default:
			return base;
	}
}

// Locate and parse the SmartArt drawing cache (diagrams/drawingN.xml).
//
// The slide-level rel with id `dataRId` points at diagrams/dataN.xml; the
// drawing cache is a separate rel on the *slide* (not the data part)
// with type DIAGRAM_DRAWING_REL_TYPE, living in the same diagrams/ dir.
// We take the first drawing rel on the slide if a specific one can't be
// paired by filename match — SmartArt frames on a slide usually don't
// share a drawing cache, and matching by data-part basename catches the
// common case.
//
// Returns a parsed DOM Document of the drawing (with dsp: elements
// already rewritten into the p-main namespace so the standard shape
// walker can consume them), or null if no drawing is available.
export async function findSmartArtDrawingDoc(
	pkg: OpenXmlPackage,
	slidePath: string,
	slideRels: Map<string, Relationship>,
	dataRId: string,
): Promise<Document | null> {
	const dataRel = slideRels.get(dataRId);
	if (!dataRel || dataRel.type !== DIAGRAM_DATA_REL_TYPE) return null;
	const dataPath = resolveRelTarget(slidePath, dataRel.target);
	// Expected pairing: data1.xml ↔ drawing1.xml (same dir, same N).
	// Rather than rely on the numeric match, scan the slide's drawing
	// rels for any that live in the same directory as the data part.
	const dataDir = dataPath.slice(0, dataPath.lastIndexOf("/") + 1);
	let drawingPath: string | null = null;
	for (const r of slideRels.values()) {
		if (r.type !== DIAGRAM_DRAWING_REL_TYPE) continue;
		const candidate = resolveRelTarget(slidePath, r.target);
		if (candidate.startsWith(dataDir)) {
			drawingPath = candidate;
			break;
		}
	}
	if (!drawingPath) return null;

	const txt = await pkg.loadText(drawingPath);
	if (!txt) return null;

	// Rewrite the dsp namespace to p-main so the existing walker matches
	// dsp:sp / dsp:spTree / dsp:spPr / dsp:nvSpPr / dsp:txBody — all of
	// which mirror the p-main element set 1:1 in structure.
	const rewritten = txt.replace(DSP_NS, P_NS);
	return pkg.parseXml(rewritten);
}

// Expand a SmartArt drawing Document into the existing shape list. Because
// the dsp namespace was rewritten into p-main, the body of the drawing
// looks like a plain slide spTree (minus cSld wrapper) and we can reuse
// the static-shapes walker by handing it the spTree directly.
export function parseSmartArtShapesFromDoc(
	doc: Document,
	ctx: SlideParseContext,
	parseShapesFromSpTree: (spTree: Element, ctx: SlideParseContext) => ShapeLike[],
): ShapeLike[] {
	// doc.documentElement is <dsp:drawing> (now in p-main). Its direct
	// child is the spTree.
	const root = doc.documentElement;
	let spTree: Element | null = null;
	for (const c of Array.from(root.children)) {
		if (c.namespaceURI === P_NS && c.localName === "spTree") {
			spTree = c;
			break;
		}
	}
	if (!spTree) return [];
	return parseShapesFromSpTree(spTree, ctx);
}
