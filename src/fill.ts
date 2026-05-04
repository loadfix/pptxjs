// Shared fill and line (stroke) resolution for shape spPr.
//
// Fill is a discriminated union:
//   SolidFill | GradientFill | BlipFill | PatternFill | NoFill
//
// `parseFillElement` dispatches on the child kind (<a:solidFill>,
// <a:gradFill>, <a:blipFill>, <a:pattFill>, <a:noFill>). The legacy
// `parseSolidFill` is retained with its narrow contract (returns SolidFill
// | null) so existing call sites still compile.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ClrMap, ThemeColors, resolveColorElement } from './theme';

export interface SolidFill {
	kind: 'solid';
	colorHex: string;
	alpha: number | null;
}

export interface NoFill {
	kind: 'none';
}

export interface GradientStop {
	posPermille: number;
	colorHex: string;
	alpha: number | null;
}

export interface GradientFill {
	kind: 'gradient';
	stops: GradientStop[];
	variant:
		| { kind: 'linear'; angle: number; scaled: boolean }
		| { kind: 'path'; path: 'circle' | 'rect' | 'shape' };
}

export interface BlipTile {
	txPermille: number;
	tyPermille: number;
	sxPermille: number;
	syPermille: number;
	flip: string | null;
	algn: string | null;
}

export interface BlipFill {
	kind: 'blip';
	rId: string;
	srcRectPermille: { l: number; t: number; r: number; b: number } | null;
	stretch: boolean;
	tile: BlipTile | null;
	alphaModFix: number | null;
}

export interface PatternFill {
	kind: 'pattern';
	preset: string;
	fgColorHex: string;
	bgColorHex: string;
}

export type Fill = SolidFill | GradientFill | BlipFill | PatternFill | NoFill;

export interface LineEnd {
	type: string;
	w: string | null;
	len: string | null;
}

export interface LineStyle {
	// EMU width (9525 EMU = 1 px at 96 DPI). null means inherit/default.
	widthEmu: number | null;
	fill: Fill | null;
	dash:
		| 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot'
		| 'dot' | 'sysDash' | 'sysDashDot' | 'sysDot' | null;
	cap: 'flat' | 'sq' | 'rnd' | null;
	join: 'round' | 'bevel' | 'miter' | null;
	cmpd: 'sng' | 'dbl' | 'thickThin' | 'thinThick' | 'tri' | null;
	headEnd: LineEnd | null;
	tailEnd: LineEnd | null;
}

// Backward-compatible solidFill parser. Existing call sites pass a
// <a:solidFill> element directly; return a SolidFill or null.
export function parseSolidFill(
	solidFill: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): SolidFill | null {
	if (!solidFill) return null;
	const colorEl = firstColorChild(solidFill);
	if (!colorEl) return null;
	const r = resolveColorElement(colorEl, clrMap, theme);
	if (!r || !r.colorHex) return null;
	return { kind: 'solid', colorHex: r.colorHex, alpha: r.alpha };
}

// Dispatch on a fill-bearing DrawingML element. Returns null if the element
// isn't one of the recognised kinds.
export function parseFillElement(
	fillEl: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): Fill | null {
	if (!fillEl) return null;
	if (fillEl.namespaceURI !== A_NS.a) return null;
	switch (fillEl.localName) {
		case 'solidFill': return parseSolidFill(fillEl, clrMap, theme);
		case 'noFill': return { kind: 'none' };
		case 'gradFill': return parseGradFill(fillEl, clrMap, theme);
		case 'blipFill': return parseBlipFill(fillEl);
		case 'pattFill': return parsePattFill(fillEl, clrMap, theme);
	}
	return null;
}

function firstColorChild(el: Element): Element | null {
	for (const c of Array.from(el.children)) {
		if (c.namespaceURI !== A_NS.a) continue;
		const ln = c.localName;
		if (
			ln === 'srgbClr' || ln === 'schemeClr' || ln === 'prstClr'
			|| ln === 'sysClr' || ln === 'scrgbClr' || ln === 'hslClr'
		) return c;
	}
	return null;
}

function parseGradFill(el: Element, clrMap: ClrMap, theme: ThemeColors): GradientFill | null {
	const gsLst = firstChildNS(el, A_NS.a, "gsLst");
	const stops: GradientStop[] = [];
	if (gsLst) {
		for (const gs of Array.from(gsLst.children)) {
			if (gs.namespaceURI !== A_NS.a || gs.localName !== 'gs') continue;
			const posAttr = gs.getAttribute("pos");
			const pos = posAttr != null ? Number(posAttr) : 0;
			const colorEl = firstColorChild(gs);
			if (!colorEl) continue;
			const resolved = resolveColorElement(colorEl, clrMap, theme);
			if (!resolved || !resolved.colorHex) continue;
			stops.push({
				posPermille: pos,
				colorHex: resolved.colorHex,
				alpha: resolved.alpha,
			});
		}
	}

	let variant: GradientFill['variant'];
	const lin = firstChildNS(el, A_NS.a, "lin");
	const path = firstChildNS(el, A_NS.a, "path");
	if (lin) {
		const ang = Number(lin.getAttribute("ang")) || 0;
		const scaled = lin.getAttribute("scaled") === "1";
		variant = { kind: 'linear', angle: ang, scaled };
	} else if (path) {
		const p = path.getAttribute("path");
		const pp: 'circle' | 'rect' | 'shape' =
			p === 'circle' || p === 'rect' || p === 'shape' ? p : 'rect';
		variant = { kind: 'path', path: pp };
	} else {
		// Default to linear horizontal if nothing specified.
		variant = { kind: 'linear', angle: 0, scaled: false };
	}

	return { kind: 'gradient', stops, variant };
}

function parseBlipFill(el: Element): BlipFill | null {
	const blip = firstChildNS(el, A_NS.a, "blip");
	const rId = blip?.getAttributeNS(A_NS.r, "embed") ?? null;
	if (!rId) return null;

	let alphaModFix: number | null = null;
	if (blip) {
		const am = firstChildNS(blip, A_NS.a, "alphaModFix");
		if (am) {
			const v = am.getAttribute("amt");
			if (v != null) {
				const n = Number(v);
				if (!Number.isNaN(n)) alphaModFix = n;
			}
		}
	}

	let srcRectPermille: BlipFill['srcRectPermille'] = null;
	const srcRect = firstChildNS(el, A_NS.a, "srcRect");
	if (srcRect) {
		srcRectPermille = {
			l: Number(srcRect.getAttribute("l")) || 0,
			t: Number(srcRect.getAttribute("t")) || 0,
			r: Number(srcRect.getAttribute("r")) || 0,
			b: Number(srcRect.getAttribute("b")) || 0,
		};
	}

	const stretchEl = firstChildNS(el, A_NS.a, "stretch");
	const tileEl = firstChildNS(el, A_NS.a, "tile");
	let stretch = true;
	let tile: BlipTile | null = null;
	if (tileEl && !stretchEl) {
		stretch = false;
		tile = {
			txPermille: Number(tileEl.getAttribute("tx")) || 0,
			tyPermille: Number(tileEl.getAttribute("ty")) || 0,
			sxPermille: Number(tileEl.getAttribute("sx")) || 100000,
			syPermille: Number(tileEl.getAttribute("sy")) || 100000,
			flip: tileEl.getAttribute("flip"),
			algn: tileEl.getAttribute("algn"),
		};
	}

	return { kind: 'blip', rId, srcRectPermille, stretch, tile, alphaModFix };
}

function parsePattFill(el: Element, clrMap: ClrMap, theme: ThemeColors): PatternFill | null {
	const preset = el.getAttribute("prst") || 'pct50';
	const fgEl = firstChildNS(el, A_NS.a, "fgClr");
	const bgEl = firstChildNS(el, A_NS.a, "bgClr");
	const fgColor = fgEl ? firstColorChild(fgEl) : null;
	const bgColor = bgEl ? firstColorChild(bgEl) : null;
	const fg = fgColor ? resolveColorElement(fgColor, clrMap, theme) : null;
	const bg = bgColor ? resolveColorElement(bgColor, clrMap, theme) : null;
	if (!fg || !bg || !fg.colorHex || !bg.colorHex) return null;
	return { kind: 'pattern', preset, fgColorHex: fg.colorHex, bgColorHex: bg.colorHex };
}

// Parse an `<a:ln>` element. `<a:noFill/>` yields a fill of kind 'none' *and*
// width zero so the renderer can suppress the border.
export function parseLine(ln: Element | null, clrMap: ClrMap, theme: ThemeColors): LineStyle | null {
	if (!ln) return null;
	const widthAttr = ln.getAttribute("w");
	const widthEmu = widthAttr != null ? Number(widthAttr) : null;

	// Find a fill child among any of the known fill kinds.
	let fill: Fill | null = null;
	for (const child of Array.from(ln.children)) {
		if (child.namespaceURI !== A_NS.a) continue;
		const ln2 = child.localName;
		if (
			ln2 === 'solidFill' || ln2 === 'noFill' || ln2 === 'gradFill'
			|| ln2 === 'blipFill' || ln2 === 'pattFill'
		) {
			fill = parseFillElement(child, clrMap, theme);
			break;
		}
	}

	const prstDash = firstChildNS(ln, A_NS.a, "prstDash");
	let dash: LineStyle['dash'] = null;
	if (prstDash) {
		const v = prstDash.getAttribute("val");
		if (
			v === 'solid' || v === 'dash' || v === 'dashDot' || v === 'lgDash'
			|| v === 'lgDashDot' || v === 'dot' || v === 'sysDash'
			|| v === 'sysDashDot' || v === 'sysDot'
		) dash = v;
	}

	const capAttr = ln.getAttribute("cap");
	const cap: LineStyle['cap'] =
		capAttr === 'flat' || capAttr === 'sq' || capAttr === 'rnd' ? capAttr : null;

	const cmpdAttr = ln.getAttribute("cmpd");
	const cmpd: LineStyle['cmpd'] =
		cmpdAttr === 'sng' || cmpdAttr === 'dbl' || cmpdAttr === 'thickThin'
			|| cmpdAttr === 'thinThick' || cmpdAttr === 'tri' ? cmpdAttr : null;

	let join: LineStyle['join'] = null;
	if (firstChildNS(ln, A_NS.a, "round")) join = 'round';
	else if (firstChildNS(ln, A_NS.a, "bevel")) join = 'bevel';
	else if (firstChildNS(ln, A_NS.a, "miter")) join = 'miter';

	const parseEnd = (e: Element | null): LineEnd | null => {
		if (!e) return null;
		const type = e.getAttribute("type") || 'none';
		return {
			type,
			w: e.getAttribute("w"),
			len: e.getAttribute("len"),
		};
	};
	const headEnd = parseEnd(firstChildNS(ln, A_NS.a, "headEnd"));
	const tailEnd = parseEnd(firstChildNS(ln, A_NS.a, "tailEnd"));

	if (
		widthEmu == null && !fill && !dash && !cap && !cmpd && !join
		&& !headEnd && !tailEnd
	) return null;

	// If fill is NoFill, zero the width so the renderer suppresses the border
	// (matching the previous behaviour where <a:noFill/> → widthEmu 0).
	const finalWidth = fill && fill.kind === 'none' ? 0 : widthEmu;
	return { widthEmu: finalWidth, fill, dash, cap, join, cmpd, headEnd, tailEnd };
}
