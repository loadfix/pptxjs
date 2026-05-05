// Parse <a:effectLst> under <p:spPr>.
//
// OOXML effect elements covered:
//   <a:outerShdw>  — outer drop shadow
//   <a:innerShdw>  — inner shadow
//   <a:glow>       — coloured glow around the shape
//   <a:softEdge>   — feathered edge (no color)
//   <a:reflection> — mirrored reflection below the shape
//   <a:blur>       — gaussian blur applied to the shape
//
// Shared conventions:
//   - blurRad, dist, rad are in EMU.
//   - dir is in 60000ths of a degree, 0 = east, increasing clockwise.
//   - stA/endA, stPos/endPos are 1000ths of a percent (per-mille).
//   - Colours are resolved via the shared resolveColorElement helper so we
//     accept srgbClr / schemeClr / prstClr / sysClr / scrgbClr / hslClr.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ClrMap, ThemeColors, resolveColorElement } from './theme';

export interface OuterShadow {
	kind: 'outerShdw';
	// All distances in EMU; caller converts to px at render time.
	blurRadEmu: number;
	distEmu: number;
	// Angle in 60000ths of a degree (OOXML convention).
	dir60k: number;
	colorHex: string;
	alpha: number | null;
	sxPermille: number;  // horizontal scale
	syPermille: number;  // vertical scale
	kxPermille: number;  // horizontal skew (10000ths of a degree in OOXML, but kept as permille here for symmetry with sx/sy)
	kyPermille: number;  // vertical skew
	algn: string | null; // alignment: "b"|"bl"|"br"|"ctr"|"l"|"r"|"t"|"tl"|"tr"
	rotWithShape: boolean;
}

export interface InnerShadow {
	kind: 'innerShdw';
	blurRadEmu: number;
	distEmu: number;
	dir60k: number;
	colorHex: string;
	alpha: number | null;
}

export interface Glow {
	kind: 'glow';
	radEmu: number;
	colorHex: string;
	alpha: number | null;
}

export interface SoftEdge {
	kind: 'softEdge';
	radEmu: number;
}

export interface Reflection {
	kind: 'reflection';
	blurRadEmu: number;
	stAPermille: number;   // start alpha
	stPosPermille: number; // start position along gradient
	endAPermille: number;  // end alpha
	endPosPermille: number;
	distEmu: number;
	dir60k: number;
	sxPermille: number;
	syPermille: number;
	kxPermille: number;
	kyPermille: number;
	algn: string | null;
	rotWithShape: boolean;
}

export interface Blur {
	kind: 'blur';
	radEmu: number;
	grow: boolean;
}

export interface ShapeEffects {
	outerShadow: OuterShadow | null;
	innerShadow: InnerShadow | null;
	glow: Glow | null;
	softEdge: SoftEdge | null;
	reflection: Reflection | null;
	blur: Blur | null;
}

export function emptyShapeEffects(): ShapeEffects {
	return {
		outerShadow: null,
		innerShadow: null,
		glow: null,
		softEdge: null,
		reflection: null,
		blur: null,
	};
}

// Resolve the first color-bearing child of an effect element.
function resolveEffectColor(el: Element, clrMap: ClrMap, theme: ThemeColors): { colorHex: string; alpha: number | null } | null {
	for (const c of Array.from(el.children)) {
		if (c.namespaceURI !== A_NS.a) continue;
		const ln = c.localName;
		if (
			ln === 'srgbClr' || ln === 'schemeClr' || ln === 'prstClr'
			|| ln === 'sysClr' || ln === 'scrgbClr' || ln === 'hslClr'
		) {
			const r = resolveColorElement(c, clrMap, theme);
			if (!r || !r.colorHex) return null;
			return { colorHex: r.colorHex, alpha: r.alpha };
		}
	}
	return null;
}

function num(el: Element, name: string, fallback: number): number {
	const v = el.getAttribute(name);
	if (v == null || v === '') return fallback;
	const n = Number(v);
	return Number.isNaN(n) ? fallback : n;
}

function bool(el: Element, name: string, fallback: boolean): boolean {
	const v = el.getAttribute(name);
	if (v == null) return fallback;
	return v === '1' || v === 'true';
}

export function parseEffectList(
	effectLstEl: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): ShapeEffects | null {
	if (!effectLstEl) return null;
	if (effectLstEl.namespaceURI !== A_NS.a || effectLstEl.localName !== 'effectLst') return null;

	const out = emptyShapeEffects();
	let found = false;

	for (const child of Array.from(effectLstEl.children)) {
		if (child.namespaceURI !== A_NS.a) continue;
		switch (child.localName) {
			case 'outerShdw': {
				const color = resolveEffectColor(child, clrMap, theme);
				if (!color) break;
				out.outerShadow = {
					kind: 'outerShdw',
					blurRadEmu: num(child, 'blurRad', 0),
					distEmu: num(child, 'dist', 0),
					dir60k: num(child, 'dir', 0),
					colorHex: color.colorHex,
					alpha: color.alpha,
					sxPermille: num(child, 'sx', 100000),
					syPermille: num(child, 'sy', 100000),
					kxPermille: num(child, 'kx', 0),
					kyPermille: num(child, 'ky', 0),
					algn: child.getAttribute('algn'),
					rotWithShape: bool(child, 'rotWithShape', true),
				};
				found = true;
				break;
			}
			case 'innerShdw': {
				const color = resolveEffectColor(child, clrMap, theme);
				if (!color) break;
				out.innerShadow = {
					kind: 'innerShdw',
					blurRadEmu: num(child, 'blurRad', 0),
					distEmu: num(child, 'dist', 0),
					dir60k: num(child, 'dir', 0),
					colorHex: color.colorHex,
					alpha: color.alpha,
				};
				found = true;
				break;
			}
			case 'glow': {
				const color = resolveEffectColor(child, clrMap, theme);
				if (!color) break;
				out.glow = {
					kind: 'glow',
					radEmu: num(child, 'rad', 0),
					colorHex: color.colorHex,
					alpha: color.alpha,
				};
				found = true;
				break;
			}
			case 'softEdge': {
				out.softEdge = {
					kind: 'softEdge',
					radEmu: num(child, 'rad', 0),
				};
				found = true;
				break;
			}
			case 'reflection': {
				out.reflection = {
					kind: 'reflection',
					blurRadEmu: num(child, 'blurRad', 0),
					stAPermille: num(child, 'stA', 100000),
					stPosPermille: num(child, 'stPos', 0),
					endAPermille: num(child, 'endA', 0),
					endPosPermille: num(child, 'endPos', 100000),
					distEmu: num(child, 'dist', 0),
					dir60k: num(child, 'dir', 5400000), // default pointing straight down
					sxPermille: num(child, 'sx', 100000),
					syPermille: num(child, 'sy', -100000),
					kxPermille: num(child, 'kx', 0),
					kyPermille: num(child, 'ky', 0),
					algn: child.getAttribute('algn'),
					rotWithShape: bool(child, 'rotWithShape', true),
				};
				found = true;
				break;
			}
			case 'blur': {
				out.blur = {
					kind: 'blur',
					radEmu: num(child, 'rad', 0),
					grow: bool(child, 'grow', true),
				};
				found = true;
				break;
			}
		}
	}

	return found ? out : null;
}

// Convenience: pull <a:effectLst> from a <p:spPr> and parse it.
export function parseEffectsFromSpPr(
	spPr: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): ShapeEffects | null {
	if (!spPr) return null;
	const effectLst = firstChildNS(spPr, A_NS.a, 'effectLst');
	return parseEffectList(effectLst, clrMap, theme);
}
