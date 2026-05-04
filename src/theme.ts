// Theme color resolution for PPTX.
//
// Chain: `<a:schemeClr val="tx1"/>`
//   → master's `<p:clrMap tx1="dk1" bg1="lt1" .../>`     (slot → scheme slot)
//   → theme's `<a:clrScheme>` children `<a:dk1>...</a:dk1>` etc.
//     (scheme slot → concrete `<a:srgbClr>` or `<a:sysClr lastClr="...">`)
//
// Modifiers on the schemeClr (tint, shade, lumMod, lumOff, satMod, alpha) are
// not yet applied — only the base color is returned. See TODO in `resolveSchemeClr`.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';

const SCHEME_SLOTS = [
	"dk1", "lt1", "dk2", "lt2",
	"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
	"hlink", "folHlink",
] as const;
type SchemeSlot = typeof SCHEME_SLOTS[number];

// Slots used on <a:schemeClr val="..."/>. bg1/tx1/bg2/tx2 are indirected
// through the master clrMap; the others map straight through.
const MAPPABLE_SLOTS = new Set(["bg1", "tx1", "bg2", "tx2"]);

export interface ThemeColors {
	scheme: Partial<Record<SchemeSlot, string>>;     // slot name → `#RRGGBB`
	// majorLatin is the `<a:majorFont>/<a:latin typeface="...">` (body headings / titles → `+mj-lt`).
	// minorLatin is the `<a:minorFont>/<a:latin typeface="...">` (body text → `+mn-lt`).
	majorLatin: string | null;
	minorLatin: string | null;
	// fmtScheme children as raw elements — resolved lazily against a phClr.
	// Indexed 1..N, matching the 1-based ECMA-376 indexing (bgRef idx=1001 → bgFillStyleLst[0]).
	bgFillStyles: Element[];
}

// Master clrMap attributes, e.g. { tx1: "dk1", bg1: "lt1", accent1: "accent1", ... }.
export type ClrMap = Record<string, string>;

export function emptyThemeColors(): ThemeColors {
	return { scheme: {}, majorLatin: null, minorLatin: null, bgFillStyles: [] };
}

// Resolve a typeface string from a run/paragraph style. Returns a concrete
// family (or null) — `+mj-lt` → theme.majorLatin, `+mn-lt` → theme.minorLatin,
// anything else passes through unchanged.
export function resolveTypeface(typeface: string | null, theme: ThemeColors): string | null {
	if (!typeface) return null;
	if (typeface === "+mj-lt") return theme.majorLatin;
	if (typeface === "+mn-lt") return theme.minorLatin;
	// Other theme references (+mj-ea, +mj-cs, etc.) aren't resolved yet —
	// leave them unset so the browser falls back to its default.
	if (typeface.startsWith("+")) return null;
	return typeface;
}

export function emptyClrMap(): ClrMap {
	return {};
}

export function parseTheme(themeDoc: Document): ThemeColors {
	const out = emptyThemeColors();
	const theme = themeDoc.documentElement;
	const elems = firstChildNS(theme, A_NS.a, "themeElements");
	if (!elems) return out;

	const fmtScheme = firstChildNS(elems, A_NS.a, "fmtScheme");
	const bgFillStyleLst = fmtScheme && firstChildNS(fmtScheme, A_NS.a, "bgFillStyleLst");
	if (bgFillStyleLst) {
		for (const c of Array.from(bgFillStyleLst.children)) {
			if (c.namespaceURI === A_NS.a) out.bgFillStyles.push(c);
		}
	}

	const fontScheme = firstChildNS(elems, A_NS.a, "fontScheme");
	if (fontScheme) {
		const major = firstChildNS(fontScheme, A_NS.a, "majorFont");
		const minor = firstChildNS(fontScheme, A_NS.a, "minorFont");
		const majorLatin = major && firstChildNS(major, A_NS.a, "latin");
		const minorLatin = minor && firstChildNS(minor, A_NS.a, "latin");
		if (majorLatin) out.majorLatin = majorLatin.getAttribute("typeface") || null;
		if (minorLatin) out.minorLatin = minorLatin.getAttribute("typeface") || null;
	}

	const clrScheme = firstChildNS(elems, A_NS.a, "clrScheme");
	if (!clrScheme) return out;

	for (const slot of SCHEME_SLOTS) {
		const slotEl = firstChildNS(clrScheme, A_NS.a, slot);
		if (!slotEl) continue;
		const srgb = firstChildNS(slotEl, A_NS.a, "srgbClr");
		if (srgb) {
			const v = srgb.getAttribute("val");
			if (v && /^[0-9a-fA-F]{6}$/.test(v)) out.scheme[slot] = `#${v}`;
			continue;
		}
		// sysClr carries a `lastClr` with the system color's resolved value.
		const sys = firstChildNS(slotEl, A_NS.a, "sysClr");
		if (sys) {
			const v = sys.getAttribute("lastClr");
			if (v && /^[0-9a-fA-F]{6}$/.test(v)) out.scheme[slot] = `#${v}`;
		}
	}
	return out;
}

export function parseClrMap(masterDoc: Document): ClrMap {
	const root = masterDoc.documentElement;
	const clrMap = firstChildNS(root, A_NS.p, "clrMap");
	const out = emptyClrMap();
	if (!clrMap) return out;
	for (const attr of Array.from(clrMap.attributes)) {
		out[attr.localName ?? attr.name] = attr.value;
	}
	return out;
}

// Resolve a <a:schemeClr val="..."/> to `#RRGGBB`, or null if it can't be
// resolved. TODO: apply child modifiers (tint, shade, lumMod, lumOff, etc.).
export function resolveSchemeClr(slot: string, clrMap: ClrMap, theme: ThemeColors): string | null {
	const resolvedSlot = MAPPABLE_SLOTS.has(slot) ? (clrMap[slot] ?? slot) : slot;
	return theme.scheme[resolvedSlot as SchemeSlot] ?? null;
}
