import type { Shape } from '../presentation-parser';
import { emuToPx, positionStyle, SVG_NS } from './geom';
import { renderParagraph, type AutoNumState } from './text';
import { solidColorFromFill } from './fill-utils';
import { presetToSvgPath } from '../preset-geom';

export function renderShape(shape: Shape, cls: string): HTMLElement {
	const el = document.createElement("div");
	el.className = `${cls}-shape`;
	Object.assign(el.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));

	// A degenerate "line" (zero cx or cy) can't render a path in zero area,
	// so fall back to the background-band treatment even if custGeom is set.
	const isDegenerateLine = shape.cx === 0 || shape.cy === 0;

	if (shape.custGeom && !isDegenerateLine) {
		el.appendChild(renderCustGeomSvg(shape));
	} else if (shape.presetGeom && !isDegenerateLine) {
		const svg = renderPresetGeomSvg(shape);
		if (svg) {
			el.appendChild(svg);
		} else {
			// Unknown preset — fall back to the plain-box look.
			applyBoxFill(el, shape);
		}
	} else {
		applyBoxFill(el, shape);
	}
	const autoNumState: AutoNumState = new Map();
	for (const p of shape.paragraphs) {
		el.appendChild(renderParagraph(p, autoNumState));
	}
	return el;
}

// Emit the fill / border as CSS on the shape's <div> (legacy path for shapes
// with no custGeom or preset, or with an unrecognised preset).
function applyBoxFill(el: HTMLElement, shape: Shape): void {
	if (shape.fill?.kind === 'solid') el.style.background = shape.fill.colorHex;
	// Wave 2 — render non-solid line fills (gradient/blip/pattern) properly.
	const lineColor = solidColorFromFill(shape.line?.fill ?? null);
	if (shape.line && lineColor) {
		const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
		const isHLine = shape.cy === 0;
		const isVLine = shape.cx === 0;
		if (isHLine || isVLine) {
			el.style.background = lineColor;
			if (isHLine) el.style.height = `${widthPx}px`;
			if (isVLine) el.style.width = `${widthPx}px`;
		} else {
			el.style.border = `${widthPx}px solid ${lineColor}`;
		}
	}
}

// Shared machinery that turns a (vbW, vbH, d) triple into an <svg> matching
// the shape's fill/stroke styling. Used by both custGeom and presetGeom.
function buildShapeSvg(shape: Shape, vbW: number, vbH: number, d: string, closed: boolean): SVGSVGElement {
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	svg.setAttribute("preserveAspectRatio", "none");
	// overflow:visible so a stroke at the edge isn't clipped by viewBox, but
	// keep the SVG sized to the parent so content can't escape the shape box.
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%" });
	const path = document.createElementNS(SVG_NS, "path");
	path.setAttribute("d", d);
	path.setAttribute("fill", shape.fill?.kind === 'solid' ? shape.fill.colorHex : "none");
	const strokeColor = solidColorFromFill(shape.line?.fill ?? null);
	if (strokeColor) {
		path.setAttribute("stroke", strokeColor);
		// With vector-effect="non-scaling-stroke" the browser interprets
		// stroke-width in screen (px) units regardless of the viewBox, so
		// convert the EMU width accordingly.
		const widthPx = shape.line!.widthEmu != null ? Math.max(emuToPx(shape.line!.widthEmu), 0.5) : 1;
		path.setAttribute("stroke-width", String(widthPx));
		path.setAttribute("vector-effect", "non-scaling-stroke");
	} else if (!closed && shape.fill?.kind !== 'solid') {
		// Open path with no fill and no line: fall back to a thin default
		// stroke so the path is at least visible.
		path.setAttribute("stroke", "#000");
		path.setAttribute("stroke-width", "1");
		path.setAttribute("vector-effect", "non-scaling-stroke");
	}
	svg.appendChild(path);
	return svg;
}

export function renderCustGeomSvg(shape: Shape): SVGSVGElement {
	// For a degenerate path (pathW or pathH is 0 — e.g. a horizontal line),
	// fall back to the shape's EMU size to avoid a zero-area viewBox.
	const vbW = shape.custGeom!.pathW || Math.max(shape.cx, 1);
	const vbH = shape.custGeom!.pathH || Math.max(shape.cy, 1);
	return buildShapeSvg(shape, vbW, vbH, shape.custGeom!.d, shape.custGeom!.closed);
}

export function renderPresetGeomSvg(shape: Shape): SVGSVGElement | null {
	const pg = shape.presetGeom!;
	const vbW = Math.max(shape.cx, 1);
	const vbH = Math.max(shape.cy, 1);
	const d = presetToSvgPath(pg.name, vbW, vbH, pg.avLst);
	if (d == null) return null;
	// Every preset in our implemented set emits a closed silhouette except
	// line / straightConnector1.
	const closed = pg.name !== "line" && pg.name !== "straightConnector1";
	return buildShapeSvg(shape, vbW, vbH, d, closed);
}
