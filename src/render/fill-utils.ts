// Extract a CSS color string from a Fill, for cases where we only know how
// to render solids today (e.g. shape borders / SVG strokes). Non-solid fills
// should use `fillToCssBackground` below.

import type { Fill, GradientFill, BlipFill, PatternFill, LineStyle } from '../fill';
import { withAlphaHex } from '../color-math';
import { SVG_NS } from './geom';

export function solidColorFromFill(fill: Fill | null): string | null {
	if (!fill) return null;
	if (fill.kind === 'solid') return withAlphaHex(fill.colorHex, fill.alpha);
	return null;
}

// A best-effort single-color approximation of any Fill, for places where the
// rendering surface can only carry a flat color (e.g. CSS `border:`). For
// non-solids: gradients return the first stop, patterns return fgColor, blips
// return null (caller supplies its own neutral). Alphas are ignored.
export function approxSolidColorFromFill(fill: Fill | null): string | null {
	if (!fill) return null;
	switch (fill.kind) {
		case 'none': return null;
		case 'solid': return fill.colorHex;
		case 'gradient': {
			if (fill.stops.length === 0) return null;
			const sorted = [...fill.stops].sort((a, b) => a.posPermille - b.posPermille);
			return sorted[0].colorHex;
		}
		case 'pattern': return fill.fgColorHex;
		case 'blip': return null;
	}
}

// Map DrawingML <a:prstDash> onto the closest CSS border-style token. "solid"
// is used as the default when dash is null/unknown. Dash-dot variants fall
// back to "dashed" since CSS has no native dash-dot border-style.
export function lineDashToCss(dash: LineStyle['dash']): string {
	switch (dash) {
		case 'dash':
		case 'lgDash':
		case 'sysDash':
		case 'dashDot':
		case 'lgDashDot':
		case 'sysDashDot':
			return 'dashed';
		case 'dot':
		case 'sysDot':
			return 'dotted';
		case 'solid':
		case null:
		case undefined:
		default:
			return 'solid';
	}
}

// Map DrawingML <a:prstDash> onto an SVG `stroke-dasharray` expressed in
// stroke-width units. Returns null for solid/missing (the caller should
// omit the attribute entirely). Patterns are rough visual approximations
// of PowerPoint's rendering — fidelity isn't pixel-perfect.
export function svgDashArray(dash: LineStyle['dash']): string | null {
	switch (dash) {
		case 'dash':       return '4 2';
		case 'dashDot':    return '4 2 1 2';
		case 'lgDash':     return '8 3';
		case 'lgDashDot':  return '8 3 1 3';
		case 'dot':        return '1 2';
		case 'sysDash':    return '3 1';
		case 'sysDashDot': return '3 1 1 1';
		case 'sysDot':     return '1 1';
		case 'solid':
		case null:
		case undefined:
		default:
			return null;
	}
}

// A monotonically-increasing counter for <defs> ids so multiple invocations
// of `fillToSvgPaint` don't collide on `url(#id)`.
let svgPaintIdCounter = 0;

// SVG paint description returned by `fillToSvgPaint`. Callers use `paint` as
// the value of stroke/fill (e.g. `#rrggbb` for solids, `url(#id)` for defs)
// and append every element in `defs` to their svg's <defs>.
export interface SvgPaint {
	paint: string;
	defs: SVGElement[];
}

// Convert any Fill into an SVG paint. Returns null when the fill is absent
// or <a:noFill/>. Solids come back with `paint = '#rrggbb'` and no defs;
// gradients/patterns/blips return a `url(#id)` paint plus the paintServer
// element that should be appended to the caller's <defs>.
export function fillToSvgPaint(
	fill: Fill | null,
	embedUrls: Map<string, string>,
	idPrefix = 'pptxjs-paint',
): SvgPaint | null {
	if (!fill) return null;
	switch (fill.kind) {
		case 'none': return null;
		case 'solid': return { paint: withAlphaHex(fill.colorHex, fill.alpha), defs: [] };
		case 'gradient': return svgPaintForGradient(fill, idPrefix);
		case 'pattern': return svgPaintForPattern(fill, idPrefix);
		case 'blip': return svgPaintForBlip(fill, embedUrls, idPrefix);
	}
}

function nextPaintId(idPrefix: string): string {
	svgPaintIdCounter += 1;
	return `${idPrefix}-${svgPaintIdCounter}`;
}

function svgPaintForGradient(fill: GradientFill, idPrefix: string): SvgPaint | null {
	if (fill.stops.length === 0) return null;
	const stops = [...fill.stops].sort((a, b) => a.posPermille - b.posPermille);
	const id = nextPaintId(idPrefix);
	let el: SVGElement;
	if (fill.variant.kind === 'linear') {
		// Convert OOXML angle (clockwise from east, 60000ths of a degree) to
		// SVG gradient vector coordinates on a unit box. The gradient vector
		// points in the direction of increasing stop position.
		const ooxmlDeg = fill.variant.angle / 60000;
		const rad = (ooxmlDeg * Math.PI) / 180;
		const dx = Math.cos(rad);
		const dy = Math.sin(rad);
		// Project a unit direction onto [0,1]x[0,1] so the vector spans the
		// bounding box diagonally when at 45°.
		const x1 = 0.5 - dx / 2;
		const y1 = 0.5 - dy / 2;
		const x2 = 0.5 + dx / 2;
		const y2 = 0.5 + dy / 2;
		el = document.createElementNS(SVG_NS, 'linearGradient');
		el.setAttribute('id', id);
		el.setAttribute('gradientUnits', 'objectBoundingBox');
		el.setAttribute('x1', x1.toFixed(4));
		el.setAttribute('y1', y1.toFixed(4));
		el.setAttribute('x2', x2.toFixed(4));
		el.setAttribute('y2', y2.toFixed(4));
	} else {
		// Path-type gradients (circle / rect / shape) — map to a centered
		// radial gradient. SVG doesn't natively render rect/shape gradients.
		el = document.createElementNS(SVG_NS, 'radialGradient');
		el.setAttribute('id', id);
		el.setAttribute('gradientUnits', 'objectBoundingBox');
		el.setAttribute('cx', '0.5');
		el.setAttribute('cy', '0.5');
		el.setAttribute('r', '0.5');
	}
	for (const s of stops) {
		const stop = document.createElementNS(SVG_NS, 'stop');
		const offset = Math.max(0, Math.min(100, s.posPermille / 1000));
		stop.setAttribute('offset', `${offset.toFixed(2)}%`);
		stop.setAttribute('stop-color', s.colorHex);
		if (s.alpha != null) {
			const a = Math.max(0, Math.min(1, s.alpha / 100000));
			stop.setAttribute('stop-opacity', a.toFixed(3));
		}
		el.appendChild(stop);
	}
	return { paint: `url(#${id})`, defs: [el] };
}

function svgPaintForPattern(fill: PatternFill, idPrefix: string): SvgPaint {
	const id = nextPaintId(idPrefix);
	const pattern = document.createElementNS(SVG_NS, 'pattern');
	pattern.setAttribute('id', id);
	pattern.setAttribute('patternUnits', 'userSpaceOnUse');
	// Dimensions come from the preset's natural tile size; use a fixed 8x8
	// canvas for percent/ltHorz/etc. presets and 16x16 for brick variants.
	const svgStr = svgForPatternPreset(fill.preset, fill.fgColorHex, fill.bgColorHex);
	if (!svgStr) {
		// Unsupported preset — fall back to a solid fg swatch as the pattern.
		pattern.setAttribute('width', '8');
		pattern.setAttribute('height', '8');
		const rect = document.createElementNS(SVG_NS, 'rect');
		rect.setAttribute('width', '8');
		rect.setAttribute('height', '8');
		rect.setAttribute('fill', fill.fgColorHex);
		pattern.appendChild(rect);
		return { paint: `url(#${id})`, defs: [pattern] };
	}
	// Parse the preset SVG string and copy its inner rect/path nodes into
	// the pattern. The outer <svg>'s width/height become the pattern's
	// tile size.
	const parser = new DOMParser();
	const doc = parser.parseFromString(svgStr, 'image/svg+xml');
	const root = doc.documentElement;
	const w = root.getAttribute('width') || '8';
	const h = root.getAttribute('height') || '8';
	pattern.setAttribute('width', w);
	pattern.setAttribute('height', h);
	for (const child of Array.from(root.childNodes)) {
		if (child.nodeType === 1) {
			// Import into the live document so it's in the correct namespace.
			pattern.appendChild(document.importNode(child, true));
		}
	}
	return { paint: `url(#${id})`, defs: [pattern] };
}

function svgPaintForBlip(
	fill: BlipFill,
	embedUrls: Map<string, string>,
	idPrefix: string,
): SvgPaint | null {
	const url = embedUrls.get(fill.rId);
	if (!url) return null;
	const id = nextPaintId(idPrefix);
	const pattern = document.createElementNS(SVG_NS, 'pattern');
	pattern.setAttribute('id', id);
	pattern.setAttribute('patternUnits', 'objectBoundingBox');
	pattern.setAttribute('width', '1');
	pattern.setAttribute('height', '1');
	const image = document.createElementNS(SVG_NS, 'image');
	image.setAttribute('href', url);
	// Legacy viewers ignore `href` and require the xlink: variant.
	image.setAttributeNS('http://www.w3.org/1999/xlink', 'xlink:href', url);
	image.setAttribute('width', '1');
	image.setAttribute('height', '1');
	image.setAttribute('preserveAspectRatio', 'none');
	pattern.appendChild(image);
	return { paint: `url(#${id})`, defs: [pattern] };
}

// Describes what `fillToCssBackground` wants its caller to do. The common
// case is `{ kind: 'css', value: '<css-value>' }` which the caller assigns
// to `element.style.background`. For blips with a crop, the fill can't be
// expressed as a plain background and the caller must overlay an <img>.
export type CssFillResult =
	| { kind: 'css'; value: string }
	| { kind: 'overlay'; src: string; srcRectPermille: { l: number; t: number; r: number; b: number } | null; tile: boolean }
	| null;

// Produce a structured description of how to paint `fill` onto an HTML box.
// Returns null when the fill is absent or explicitly <a:noFill/>. Solid /
// gradient / pattern fills come back as `{ kind: 'css' }`. Blip fills come
// back as `{ kind: 'overlay' }` so the caller can honour <a:srcRect> via an
// inner <img> + overflow:hidden (plain `background-image` can't crop to a
// sub-rectangle).
export function fillToCssBackground(
	fill: Fill | null,
	embedUrls: Map<string, string>,
): CssFillResult {
	if (!fill) return null;
	switch (fill.kind) {
		case 'none': return null;
		case 'solid': return { kind: 'css', value: withAlphaHex(fill.colorHex, fill.alpha) };
		case 'gradient': {
			const v = gradientToCss(fill);
			return v ? { kind: 'css', value: v } : null;
		}
		case 'blip': {
			const url = embedUrls.get(fill.rId);
			if (!url) return null;
			return {
				kind: 'overlay',
				src: url,
				srcRectPermille: fill.srcRectPermille,
				tile: !fill.stretch,
			};
		}
		case 'pattern': {
			const v = patternToCss(fill);
			return v ? { kind: 'css', value: v } : null;
		}
	}
}

function gradientToCss(fill: GradientFill): string | null {
	if (fill.stops.length === 0) return null;
	// Sort stops by position so the CSS gradient is monotonic.
	const stops = [...fill.stops].sort((a, b) => a.posPermille - b.posPermille);
	const stopStr = stops
		.map(s => {
			const pct = Math.max(0, Math.min(100, s.posPermille / 1000));
			const color = withAlphaHex(s.colorHex, s.alpha);
			return `${color} ${pct.toFixed(2)}%`;
		})
		.join(", ");

	if (fill.variant.kind === 'linear') {
		// OOXML `ang` is in 60000ths of a degree, measured clockwise from
		// 3 o'clock (east). CSS `linear-gradient(Ndeg, ...)` measures N
		// clockwise from 12 o'clock (north) with 0deg meaning "to top".
		// To convert: cssDeg = ooxmlDeg + 90.
		const ooxmlDeg = fill.variant.angle / 60000;
		const cssDeg = (ooxmlDeg + 90) % 360;
		return `linear-gradient(${cssDeg.toFixed(2)}deg, ${stopStr})`;
	}
	// Path-type gradients (circle / rect / shape) — map everything to a
	// centered circular radial gradient; pptx has no exact CSS analogue
	// for rect/shape.
	return `radial-gradient(circle at center, ${stopStr})`;
}

function patternToCss(fill: PatternFill): string | null {
	const svg = svgForPatternPreset(fill.preset, fill.fgColorHex, fill.bgColorHex);
	if (svg) {
		// Percent-encode only the characters that matter inside a data: URL
		// value (quotes, #, newlines).
		const encoded = svg
			.replace(/\n/g, "")
			.replace(/#/g, "%23")
			.replace(/"/g, "'");
		return `url("data:image/svg+xml;utf8,${encoded}")`;
	}
	// Unsupported preset — fall back to the foreground color.
	return fill.fgColorHex;
}

// Per-mille denominator shared with pic.ts (see PERMILLE there). 100000 = 100%.
const PERMILLE = 100000;

// Shared crop-overlay implementation used by both shape.ts (for blipFill
// backgrounds) and pic.ts (for direct <p:pic> rendering). Given a wrapper
// element, an image URL, and an optional srcRect, appends an <img> to the
// wrapper that honours the crop — either by positioning an oversized <img>
// inside `overflow:hidden` (when srcRect is present) or by stretching to
// 100% otherwise. Returns the appended <img> so callers can further tweak
// it (filter/opacity/alt/title).
export function applyCropOverlay(
	wrap: HTMLElement,
	src: string,
	srcRectPermille: { l: number; t: number; r: number; b: number } | null,
): HTMLImageElement {
	const img = document.createElement('img');
	img.src = src;
	if (srcRectPermille) {
		const { l, t, r, b } = srcRectPermille;
		const cropWFrac = 1 - (l + r) / PERMILLE;
		const cropHFrac = 1 - (t + b) / PERMILLE;
		if (cropWFrac > 0 && cropHFrac > 0) {
			wrap.style.overflow = 'hidden';
			// The wrapper needs a positioning context so the <img>'s absolute
			// positioning resolves against it. Callers already set position
			// elsewhere (shape/pic wrappers are `position: absolute`), but
			// make it explicit in case the wrapper is a raw <div>.
			if (!wrap.style.position) wrap.style.position = 'relative';
			img.style.position = 'absolute';
			img.style.width = `${100 / cropWFrac}%`;
			img.style.height = `${100 / cropHFrac}%`;
			img.style.left = `${-(l / PERMILLE) * (100 / cropWFrac)}%`;
			img.style.top = `${-(t / PERMILLE) * (100 / cropHFrac)}%`;
			img.style.maxWidth = 'none';
			img.style.maxHeight = 'none';
			wrap.appendChild(img);
			return img;
		}
	}
	img.style.width = '100%';
	img.style.height = '100%';
	img.style.objectFit = 'fill';
	wrap.appendChild(img);
	return img;
}

// Build an inline SVG for a subset of OOXML pattern presets. Returns null
// when the preset isn't supported so the caller can fall back.
function svgForPatternPreset(preset: string, fg: string, bg: string): string | null {
	// Percent-based presets: density of fg dots on a bg tile.
	const pct = (coveragePct: number) => {
		// 8x8 tile; place `coveragePct` worth of fg pixels deterministically.
		const cells = 64;
		const fill = Math.round((coveragePct / 100) * cells);
		let dots = "";
		let placed = 0;
		for (let i = 0; i < cells && placed < fill; i++) {
			// Interleaved positions so small percentages stay visible.
			const x = (i * 5) % 8;
			const y = Math.floor((i * 5) / 8) % 8;
			dots += `<rect x='${x}' y='${y}' width='1' height='1' fill='${fg}'/>`;
			placed++;
		}
		return `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8' shape-rendering='crispEdges'><rect width='8' height='8' fill='${bg}'/>${dots}</svg>`;
	};

	switch (preset) {
		case 'pct5':  return pct(5);
		case 'pct10': return pct(10);
		case 'pct20': return pct(20);
		case 'pct25': return pct(25);
		case 'pct50': return pct(50);
		case 'pct75': return pct(75);

		case 'ltHorz':
			// Light horizontal stripes: 1px fg, 3px bg.
			return `<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4' shape-rendering='crispEdges'><rect width='4' height='4' fill='${bg}'/><rect y='0' width='4' height='1' fill='${fg}'/></svg>`;
		case 'ltVert':
			return `<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4' shape-rendering='crispEdges'><rect width='4' height='4' fill='${bg}'/><rect x='0' width='1' height='4' fill='${fg}'/></svg>`;
		case 'dkHorz':
			// Dense horizontal stripes: 2px fg, 2px bg.
			return `<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4' shape-rendering='crispEdges'><rect width='4' height='4' fill='${bg}'/><rect width='4' height='2' fill='${fg}'/></svg>`;
		case 'dkVert':
			return `<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4' shape-rendering='crispEdges'><rect width='4' height='4' fill='${bg}'/><rect width='2' height='4' fill='${fg}'/></svg>`;

		case 'horzBrick':
			return `<svg xmlns='http://www.w3.org/2000/svg' width='16' height='8' shape-rendering='crispEdges'><rect width='16' height='8' fill='${bg}'/><path d='M0 0 H16 M0 4 H16 M4 0 V4 M12 4 V8' stroke='${fg}' stroke-width='1'/></svg>`;
		case 'diagBrick':
			return `<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' shape-rendering='crispEdges'><rect width='16' height='16' fill='${bg}'/><path d='M0 8 L8 0 L16 8 L8 16 Z' fill='none' stroke='${fg}' stroke-width='1'/></svg>`;

		case 'plaid':
			return `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8' shape-rendering='crispEdges'><rect width='8' height='8' fill='${bg}'/><rect width='4' height='4' fill='${fg}' fill-opacity='0.5'/><rect x='4' y='4' width='4' height='4' fill='${fg}' fill-opacity='0.5'/></svg>`;

		case 'sphere':
			// Pre-rasterising a radial sphere pattern is awkward; draw a
			// filled circle on bg as a crude substitute.
			return `<svg xmlns='http://www.w3.org/2000/svg' width='8' height='8'><rect width='8' height='8' fill='${bg}'/><circle cx='4' cy='4' r='3' fill='${fg}'/></svg>`;
	}
	return null;
}

