// Slide background resolution.
//
// OOXML background chain: slide → layout → master. First hit wins. A slide
// element carries `<p:cSld><p:bg>` which may be:
//   <p:bgPr>     — direct fill (solidFill / blipFill / gradFill / pattFill)
//   <p:bgRef>    — theme reference by fmtScheme index (not resolved here)
//
// Only solidFill is resolved — either via direct <a:srgbClr> or via
// <a:schemeClr> through the slide's clrMap + theme. gradFill/blipFill/
// bgRef are parsed-through but return null.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ClrMap, ThemeColors, resolveSchemeClr } from './theme';
import { applyMods, parseColorMods } from './color-math';

export interface BackgroundFill {
	// The only fill kind currently resolved into a concrete CSS value.
	kind: 'solid';
	colorHex: string;
}

export function parseSlideBackground(
	slideDoc: Document,
	clrMap: ClrMap,
	theme: ThemeColors,
): BackgroundFill | null {
	const cSld = firstChildNS(slideDoc.documentElement, A_NS.p, "cSld");
	const bg = firstChildNS(cSld, A_NS.p, "bg");
	if (!bg) return null;

	const bgPr = firstChildNS(bg, A_NS.p, "bgPr");
	if (bgPr) {
		return resolveSolidFill(firstChildNS(bgPr, A_NS.a, "solidFill"), clrMap, theme);
	}

	const bgRef = firstChildNS(bg, A_NS.p, "bgRef");
	if (bgRef) return resolveBgRef(bgRef, clrMap, theme);

	return null;
}

// Resolve <p:bgRef idx="N"> against theme bgFillStyleLst, substituting phClr
// with the bgRef's own <a:schemeClr>/<a:srgbClr>. Only solidFill entries in
// the style list are materialised — gradFill/blipFill/pattFill return null.
function resolveBgRef(bgRef: Element, clrMap: ClrMap, theme: ThemeColors): BackgroundFill | null {
	const idxAttr = bgRef.getAttribute("idx");
	if (!idxAttr) return null;
	const idx = Number(idxAttr);
	// Index 0 or 1000 == "no fill". 1001..1999 = bgFillStyleLst; 1..999 = fillStyleLst (not used here).
	if (idx < 1001) return null;
	const styleIdx = idx - 1001;
	const styleEl = theme.bgFillStyles[styleIdx];
	if (!styleEl) return null;
	// We only know how to render solidFill entries as CSS colors.
	if (styleEl.localName !== "solidFill") return null;

	// Resolve the "placeholder color" (phClr) — the color the bgRef provides.
	const phHex = resolvePhColor(bgRef, clrMap, theme);
	if (!phHex) return null;

	// Walk the solidFill children to find the color and apply any modifiers
	// that sit on the phClr reference inside the style.
	const clr = firstChildNS(styleEl, A_NS.a, "schemeClr")
		?? firstChildNS(styleEl, A_NS.a, "srgbClr");
	if (!clr) return null;
	const slot = clr.getAttribute("val");
	// If the style's color is literal phClr, substitute; otherwise treat it
	// as a concrete color (rare for bg styles).
	let hex: string;
	if (clr.localName === "schemeClr" && slot === "phClr") {
		hex = phHex;
	} else if (clr.localName === "schemeClr" && slot) {
		const base = resolveSchemeClr(slot, clrMap, theme);
		if (!base) return null;
		hex = base;
	} else if (clr.localName === "srgbClr" && slot && /^[0-9a-fA-F]{6}$/.test(slot)) {
		hex = `#${slot}`;
	} else {
		return null;
	}
	hex = applyMods(hex, parseColorMods(clr));
	return { kind: 'solid', colorHex: hex };
}

function resolvePhColor(bgRef: Element, clrMap: ClrMap, theme: ThemeColors): string | null {
	const scheme = firstChildNS(bgRef, A_NS.a, "schemeClr");
	if (scheme) {
		const slot = scheme.getAttribute("val");
		if (!slot) return null;
		const base = resolveSchemeClr(slot, clrMap, theme);
		if (!base) return null;
		return applyMods(base, parseColorMods(scheme));
	}
	const srgb = firstChildNS(bgRef, A_NS.a, "srgbClr");
	if (srgb) {
		const v = srgb.getAttribute("val");
		if (v && /^[0-9a-fA-F]{6}$/.test(v)) return applyMods(`#${v}`, parseColorMods(srgb));
	}
	return null;
}

function resolveSolidFill(
	solidFill: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): BackgroundFill | null {
	if (!solidFill) return null;
	const srgb = firstChildNS(solidFill, A_NS.a, "srgbClr");
	if (srgb) {
		const v = srgb.getAttribute("val");
		if (v && /^[0-9a-fA-F]{6}$/.test(v)) return { kind: 'solid', colorHex: `#${v}` };
	}
	const scheme = firstChildNS(solidFill, A_NS.a, "schemeClr");
	if (scheme) {
		const slot = scheme.getAttribute("val");
		if (slot) {
			const hex = resolveSchemeClr(slot, clrMap, theme);
			if (hex) return { kind: 'solid', colorHex: hex };
		}
	}
	return null;
}
