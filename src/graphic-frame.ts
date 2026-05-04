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
