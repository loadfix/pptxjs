// Color modifier math for OOXML schemeClr/srgbClr children.
//
// Supported child elements (values are per-mille, i.e. 75000 = 75%):
//   <a:tint val="N"/>     lightens toward white.  Lmod = L * (N/100000) + (1 - N/100000)
//   <a:shade val="N"/>    darkens toward black.   Lmod = L * (N/100000)
//   <a:lumMod val="N"/>   multiplies luminance.   Lmod = L * (N/100000)
//   <a:lumOff val="N"/>   adds to luminance.      Lmod = L + (N/100000)
//   <a:alpha val="N"/>    opacity — captured by theme.extractAlpha as a 0..1
//                         value and threaded through ResolvedColor.alpha so
//                         renderers can fold it into `#RRGGBBAA` (CSS 8-hex)
//                         or rgba() output.
//
// All adjustments happen in HSL space, operating on the L channel only, and
// are applied in document order. References: ECMA-376 §20.1.2.3 and the
// OfficeOpenXML tint/shade formula used by PowerPoint.

import { A_NS } from './namespaces';

export interface ColorMods {
	// Ordered list of ops to apply, in the order they appear in the XML.
	ops: ColorModOp[];
}

type ColorModOp =
	| { kind: 'tint'; v: number }
	| { kind: 'shade'; v: number }
	| { kind: 'lumMod'; v: number }
	| { kind: 'lumOff'; v: number };

// Parse modifier children of a <a:schemeClr> or <a:srgbClr> element.
export function parseColorMods(colorEl: Element | null): ColorMods {
	const ops: ColorModOp[] = [];
	if (!colorEl) return { ops };
	for (const c of Array.from(colorEl.children)) {
		if (c.namespaceURI !== A_NS.a) continue;
		const v = c.getAttribute("val");
		if (v == null) continue;
		const num = Number(v) / 100000;
		if (Number.isNaN(num)) continue;
		if (c.localName === "tint") ops.push({ kind: 'tint', v: num });
		else if (c.localName === "shade") ops.push({ kind: 'shade', v: num });
		else if (c.localName === "lumMod") ops.push({ kind: 'lumMod', v: num });
		else if (c.localName === "lumOff") ops.push({ kind: 'lumOff', v: num });
	}
	return { ops };
}

export function applyMods(hex: string, mods: ColorMods): string {
	if (mods.ops.length === 0) return hex;
	let { h, s, l } = hexToHsl(hex);
	for (const op of mods.ops) {
		if (op.kind === 'tint') l = l * op.v + (1 - op.v);
		else if (op.kind === 'shade') l = l * op.v;
		else if (op.kind === 'lumMod') l = l * op.v;
		else if (op.kind === 'lumOff') l = l + op.v;
	}
	if (l < 0) l = 0;
	if (l > 1) l = 1;
	return hslToHex(h, s, l);
}

function hexToHsl(hex: string): { h: number; s: number; l: number } {
	const m = /^#?([0-9a-f]{6})$/i.exec(hex);
	if (!m) return { h: 0, s: 0, l: 0 };
	const n = parseInt(m[1], 16);
	const r = ((n >> 16) & 0xff) / 255;
	const g = ((n >> 8) & 0xff) / 255;
	const b = (n & 0xff) / 255;
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	if (max === min) return { h: 0, s: 0, l };
	const d = max - min;
	const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
	let h: number;
	if (max === r) h = ((g - b) / d + (g < b ? 6 : 0));
	else if (max === g) h = ((b - r) / d + 2);
	else h = ((r - g) / d + 4);
	return { h: h / 6, s, l };
}

function hslToHex(h: number, s: number, l: number): string {
	if (s === 0) {
		const v = Math.round(l * 255);
		return `#${toHex(v)}${toHex(v)}${toHex(v)}`;
	}
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	const r = hue2rgb(p, q, h + 1 / 3);
	const g = hue2rgb(p, q, h);
	const b = hue2rgb(p, q, h - 1 / 3);
	return `#${toHex(Math.round(r * 255))}${toHex(Math.round(g * 255))}${toHex(Math.round(b * 255))}`;
}

function hue2rgb(p: number, q: number, t: number): number {
	if (t < 0) t += 1;
	if (t > 1) t -= 1;
	if (t < 1 / 6) return p + (q - p) * 6 * t;
	if (t < 1 / 2) return q;
	if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
	return p;
}

function toHex(n: number): string {
	return n.toString(16).padStart(2, "0");
}

// custGeom path `fill="darken"` — one luminance step darker than the base
// color, matching PowerPoint's interpretation (equivalent to <a:lumMod
// val="75000"/>). Used when a path inside a custGeom wants to render in a
// shaded variant of the shape's own fill.
export function darken(hex: string): string {
	return applyMods(hex, { ops: [{ kind: 'lumMod', v: 0.75 }] });
}

// custGeom path `fill="lighten"` — one luminance step lighter than the base
// color (equivalent to <a:lumMod val="75000"/><a:lumOff val="25000"/>, i.e.
// l' = l * 0.75 + 0.25). Mirrors PowerPoint's "1-step lighter" behaviour.
export function lighten(hex: string): string {
	return applyMods(hex, { ops: [{ kind: 'lumMod', v: 0.75 }, { kind: 'lumOff', v: 0.25 }] });
}

// Compose a CSS color that carries the given alpha (0..1). Input is a
// `#RRGGBB` hex color (any other form, including an already-alpha'd value,
// is returned unchanged). Output is `#RRGGBBAA` — CSS Color Level 4 8-hex,
// which is supported across every major browser and is more compact than
// `rgba()`. When alpha is null, undefined, or >= 1 the input is returned
// verbatim so opaque colors don't pick up a spurious suffix.
export function withAlphaHex(hex: string, alpha: number | null | undefined): string {
	if (alpha == null || alpha >= 1) return hex;
	if (!/^#[0-9a-fA-F]{6}$/.test(hex)) return hex;
	const a = Math.max(0, Math.min(1, alpha));
	const aByte = Math.round(a * 255);
	return `${hex}${toHex(aByte)}`;
}
