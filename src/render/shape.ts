import type { Shape } from '../presentation-parser';
import type { ShapeEffects, OuterShadow, InnerShadow, Glow, SoftEdge, Reflection, Blur } from '../effects';
import { emuToPx, positionStyle, SVG_NS } from './geom';
import { renderParagraph, type AutoNumState } from './text';
import { solidColorFromFill } from './fill-utils';

export function renderShape(shape: Shape, cls: string): HTMLElement {
	const el = document.createElement("div");
	el.className = `${cls}-shape`;
	Object.assign(el.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));

	// A degenerate "line" (zero cx or cy) can't render a path in zero area,
	// so fall back to the background-band treatment even if custGeom is set.
	const isDegenerateLine = shape.cx === 0 || shape.cy === 0;

	if (shape.custGeom && !isDegenerateLine) {
		el.appendChild(renderCustGeomSvg(shape));
	} else {
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

	if (shape.effects) applyEffects(el, shape.effects);

	const autoNumState: AutoNumState = new Map();
	for (const p of shape.paragraphs) {
		el.appendChild(renderParagraph(p, autoNumState));
	}
	return el;
}

// Map a parsed <a:effectLst> onto CSS on the outer wrapper div. Approximations
// called out inline; these are visual hints, not pixel-perfect reproductions
// of what PowerPoint draws.
function applyEffects(el: HTMLElement, fx: ShapeEffects): void {
	const boxShadows: string[] = [];
	const filters: string[] = [];

	if (fx.outerShadow) boxShadows.push(outerShadowToBoxShadow(fx.outerShadow));
	if (fx.innerShadow) boxShadows.push(innerShadowToBoxShadow(fx.innerShadow));
	if (fx.glow) {
		// Glow — render as an outward box-shadow with zero offset; this keeps
		// the glow out of the `filter` chain so it doesn't interact with
		// blur/softEdge. `drop-shadow` would be more physically accurate but
		// doesn't stack as readably for rectangles.
		boxShadows.push(glowToBoxShadow(fx.glow));
	}

	// `filter` stacks blur and softEdge (both blur the shape itself).
	if (fx.blur) filters.push(blurToFilter(fx.blur));
	if (fx.softEdge) filters.push(softEdgeToFilter(fx.softEdge));

	if (boxShadows.length) el.style.boxShadow = boxShadows.join(', ');
	if (filters.length) el.style.filter = filters.join(' ');

	// Reflection — `-webkit-box-reflect` is non-standard but widely supported
	// in Chromium/WebKit browsers and gracefully ignored elsewhere. Firefox
	// has no equivalent.
	if (fx.reflection) {
		const r = reflectionToWebkitBoxReflect(fx.reflection);
		if (r) (el.style as CSSStyleDeclaration & { webkitBoxReflect?: string }).webkitBoxReflect = r;
	}
}

// Convert a direction in 60000ths of a degree + distance in EMU into CSS
// offset pixels (x, y). OOXML uses 0° = east, angles increase clockwise (y down).
function dirDistToOffsetPx(dir60k: number, distEmu: number): { dx: number; dy: number } {
	const rad = (dir60k / 60000) * Math.PI / 180;
	const distPx = emuToPx(distEmu);
	return {
		dx: distPx * Math.cos(rad),
		dy: distPx * Math.sin(rad),
	};
}

function withAlpha(hex: string, alpha: number | null): string {
	if (alpha == null || alpha >= 1) return hex;
	// #RRGGBB → rgba(r,g,b,a). alpha is already in 0..1.
	const m = /^#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/.exec(hex);
	if (!m) return hex;
	const r = parseInt(m[1], 16);
	const g = parseInt(m[2], 16);
	const b = parseInt(m[3], 16);
	return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha))})`;
}

function outerShadowToBoxShadow(s: OuterShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlpha(s.colorHex, s.alpha)}`;
}

function innerShadowToBoxShadow(s: InnerShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `inset ${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlpha(s.colorHex, s.alpha)}`;
}

function glowToBoxShadow(g: Glow): string {
	const radPx = emuToPx(g.radEmu);
	return `0 0 ${radPx.toFixed(2)}px ${radPx.toFixed(2)}px ${withAlpha(g.colorHex, g.alpha)}`;
}

function blurToFilter(b: Blur): string {
	const px = emuToPx(b.radEmu);
	return `blur(${px.toFixed(2)}px)`;
}

// softEdge is defined as a feathered alpha-mask on the shape edges. CSS has
// no direct equivalent, so we approximate with `filter: blur()` — this also
// blurs the content inside the shape, which isn't strictly correct but keeps
// the implementation simple. A future refinement could use an SVG feGaussianBlur
// + feComposite alpha mask.
function softEdgeToFilter(s: SoftEdge): string {
	const px = emuToPx(s.radEmu);
	return `blur(${px.toFixed(2)}px)`;
}

// Map <a:reflection> to `-webkit-box-reflect`. The OOXML model is far richer
// (gradient alpha stops, skew, custom offset direction); we collapse it to a
// below-reflection with an approximate gap and a linear fade.
function reflectionToWebkitBoxReflect(r: Reflection): string | null {
	const gapPx = emuToPx(r.distEmu);
	// Treat stA/endA as 0..100000 per-mille; convert to 0..1.
	const stA = Math.max(0, Math.min(1, r.stAPermille / 100000));
	const endA = Math.max(0, Math.min(1, r.endAPermille / 100000));
	// -webkit-box-reflect uses a mask-image linear-gradient; the gradient
	// goes from the near edge (full opacity / matches stA) to the far edge
	// (matches endA).
	const mask = `linear-gradient(to bottom, rgba(0,0,0,${stA.toFixed(3)}), rgba(0,0,0,${endA.toFixed(3)}))`;
	return `below ${gapPx.toFixed(2)}px ${mask}`;
}

export function renderCustGeomSvg(shape: Shape): SVGSVGElement {
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	// For a degenerate path (pathW or pathH is 0 — e.g. a horizontal line),
	// fall back to the shape's EMU size to avoid a zero-area viewBox.
	const vbW = shape.custGeom!.pathW || Math.max(shape.cx, 1);
	const vbH = shape.custGeom!.pathH || Math.max(shape.cy, 1);
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	svg.setAttribute("preserveAspectRatio", "none");
	// overflow:visible so a stroke at the edge isn't clipped by viewBox, but
	// keep the SVG sized to the parent so content can't escape the shape box.
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%" });
	const path = document.createElementNS(SVG_NS, "path");
	path.setAttribute("d", shape.custGeom!.d);
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
	} else if (!shape.custGeom!.closed && shape.fill?.kind !== 'solid') {
		// Open path with no fill and no line: fall back to a thin default
		// stroke so the path is at least visible.
		path.setAttribute("stroke", "#000");
		path.setAttribute("stroke-width", "1");
		path.setAttribute("vector-effect", "non-scaling-stroke");
	}
	svg.appendChild(path);
	return svg;
}
