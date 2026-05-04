// Shared fill and line (stroke) resolution for shape spPr.
//
// Only solidFill is resolved into a concrete CSS color — via <a:srgbClr> or
// <a:schemeClr>. gradFill / blipFill / pattFill are parsed-through and
// return null for now.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ClrMap, ThemeColors, resolveSchemeClr } from './theme';
import { applyMods, parseColorMods } from './color-math';

export interface SolidFill {
	kind: 'solid';
	colorHex: string;
}

export interface LineStyle {
	// EMU width (9525 EMU = 1 px at 96 DPI). null means inherit/default.
	widthEmu: number | null;
	fill: SolidFill | null;
}

export function parseSolidFill(
	solidFill: Element | null,
	clrMap: ClrMap,
	theme: ThemeColors,
): SolidFill | null {
	if (!solidFill) return null;
	const srgb = firstChildNS(solidFill, A_NS.a, "srgbClr");
	if (srgb) {
		const v = srgb.getAttribute("val");
		if (v && /^[0-9a-fA-F]{6}$/.test(v)) {
			return { kind: 'solid', colorHex: applyMods(`#${v}`, parseColorMods(srgb)) };
		}
	}
	const scheme = firstChildNS(solidFill, A_NS.a, "schemeClr");
	if (scheme) {
		const slot = scheme.getAttribute("val");
		if (slot) {
			const base = resolveSchemeClr(slot, clrMap, theme);
			if (base) return { kind: 'solid', colorHex: applyMods(base, parseColorMods(scheme)) };
		}
	}
	return null;
}

// Parse an `<a:ln>` element. `<a:noFill/>` yields a fill of null *and* width
// zero so the renderer can suppress the border.
export function parseLine(ln: Element | null, clrMap: ClrMap, theme: ThemeColors): LineStyle | null {
	if (!ln) return null;
	const widthAttr = ln.getAttribute("w");
	const widthEmu = widthAttr != null ? Number(widthAttr) : null;
	if (firstChildNS(ln, A_NS.a, "noFill")) {
		return { widthEmu: 0, fill: null };
	}
	const fill = parseSolidFill(firstChildNS(ln, A_NS.a, "solidFill"), clrMap, theme);
	if (widthEmu == null && !fill) return null;
	return { widthEmu, fill };
}
