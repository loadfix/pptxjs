// Extract a CSS color string from a Fill, for cases where we only know how
// to render solids today (e.g. shape borders / SVG strokes). Non-solid fills
// should use `fillToCssBackground` below.

import type { Fill, GradientFill, BlipFill, PatternFill, LineStyle } from '../fill';
import { withAlphaHex } from '../color-math';

export function solidColorFromFill(fill: Fill | null): string | null {
	if (!fill) return null;
	if (fill.kind === 'solid') return withAlphaHex(fill.colorHex, fill.alpha);
	return null;
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

// Produce a CSS value suitable for `element.style.background`. Returns null
// when the fill is absent or explicitly <a:noFill/>. Non-solid fills are
// mapped as follows:
//   gradient linear → `linear-gradient(Ndeg, stop%, ...)`
//   gradient path   → `radial-gradient(circle at center, stops)`
//   blip            → `url(<blob-url>)` with best-effort size/repeat
//   pattern         → inline SVG data URL for a small set of presets,
//                     falling back to the fg color for unsupported ones.
//
// `embedUrls` maps the rId on a BlipFill to a resolved blob/data URL. A
// BlipFill whose rId isn't in the map returns null (caller should fall back
// to whatever background already exists).
export function fillToCssBackground(
	fill: Fill | null,
	embedUrls: Map<string, string>,
): string | null {
	if (!fill) return null;
	switch (fill.kind) {
		case 'none': return null;
		case 'solid': return withAlphaHex(fill.colorHex, fill.alpha);
		case 'gradient': return gradientToCss(fill);
		case 'blip': return blipToCss(fill, embedUrls);
		case 'pattern': return patternToCss(fill);
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

function blipToCss(fill: BlipFill, embedUrls: Map<string, string>): string | null {
	const url = embedUrls.get(fill.rId);
	if (!url) return null;
	// OOXML `srcRect` crops the source image before layout; we can't truly
	// crop with just `background-image`, so we degrade to best-effort
	// size/position. Full crop fidelity would need an overlay/clip path.
	// TODO: honour srcRectPermille properly via an inner <img> + clip.
	const urlPart = `url("${url.replace(/"/g, '\\"')}")`;
	if (fill.stretch) {
		return `${urlPart} center / 100% 100% no-repeat`;
	}
	// Tile: let the browser repeat at natural size.
	return `${urlPart} top left / auto repeat`;
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

