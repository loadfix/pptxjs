// Theme color resolution for PPTX.
//
// Chain: `<a:schemeClr val="tx1"/>`
//   → master's `<p:clrMap tx1="dk1" bg1="lt1" .../>`     (slot → scheme slot)
//   → theme's `<a:clrScheme>` children `<a:dk1>...</a:dk1>` etc.
//     (scheme slot → concrete `<a:srgbClr>` or `<a:sysClr lastClr="...">`)
//
// Modifiers on the schemeClr (tint, shade, lumMod, lumOff, satMod, alpha) are
// applied via `applyMods`.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ColorMods, applyMods, parseColorMods } from './color-math';

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
		const sys = firstChildNS(slotEl, A_NS.a, "sysClr");
		if (sys) {
			const resolved = resolveSysClr(sys);
			if (resolved) out.scheme[slot] = resolved;
		}
	}
	return out;
}

// ECMA-376 §20.1.10.58 ST_SystemColorVal — map a <a:sysClr val="..."/> name to
// its canonical Windows system color. PowerPoint writes a `lastClr` attribute
// as a cache of the last-seen resolved value at save time; we use the spec
// table first and fall back to `lastClr` so older / unusual names still
// resolve. Keys are the exact `val` tokens (camelCase).
const SYS_COLOR_TABLE: Record<string, string> = {
	scrollBar: '#C8C8C8',
	background: '#FFFFFF',
	activeCaption: '#99B4D1',
	inactiveCaption: '#BFCDDB',
	menu: '#F0F0F0',
	window: '#FFFFFF',
	windowFrame: '#646464',
	menuText: '#000000',
	windowText: '#000000',
	captionText: '#000000',
	activeBorder: '#B4B4B4',
	inactiveBorder: '#F4F7FC',
	appWorkspace: '#ABABAB',
	highlight: '#3399FF',
	highlightText: '#FFFFFF',
	btnFace: '#F0F0F0',
	btnShadow: '#A0A0A0',
	grayText: '#6D6D6D',
	btnText: '#000000',
	inactiveCaptionText: '#000000',
	btnHighlight: '#FFFFFF',
	'3dDkShadow': '#696969',
	'3dLight': '#E3E3E3',
	infoText: '#000000',
	infoBk: '#FFFFE1',
	hotLight: '#0066CC',
	gradientActiveCaption: '#B9D1EA',
	gradientInactiveCaption: '#D7E4F2',
	menuHighlight: '#3399FF',
	menuBar: '#F0F0F0',
	// Legacy / alternative spellings. `3dFace` / `3dShadow` / `3dHighlight`
	// aren't part of ST_SystemColorVal but appear in the wild.
	'3dFace': '#F0F0F0',
	'3dShadow': '#A0A0A0',
	'3dHighlight': '#FFFFFF',
};

// Resolve an <a:sysClr val="..." lastClr="..."/> to `#RRGGBB`. Prefers the
// ECMA-376 table; falls back to `lastClr` when `val` is unknown. Returns
// null if neither source produces a usable color.
function resolveSysClr(el: Element): string | null {
	const val = el.getAttribute("val");
	if (val && SYS_COLOR_TABLE[val]) return SYS_COLOR_TABLE[val];
	const last = el.getAttribute("lastClr");
	if (last && /^[0-9a-fA-F]{6}$/.test(last)) return `#${last}`;
	return null;
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

// Parse <p:clrMapOvr> from a slide (<p:sld>) or slide layout (<p:sldLayout>) doc.
//
//   <p:clrMapOvr>
//     <a:masterClrMapping/>               → inherit the master's clrMap
//     <a:overrideClrMapping bg1="lt1" …/> → explicit per-part map
//   </p:clrMapOvr>
//
// Returns null to signal "inherit from the next level up". A present
// <a:overrideClrMapping> returns a freshly parsed ClrMap parallel to
// parseClrMap. Absence of <p:clrMapOvr> also returns null (same inherit
// behaviour as <a:masterClrMapping/>).
export function parseClrMapOvr(doc: Document): ClrMap | null {
	const root = doc.documentElement;
	const clrMapOvr = firstChildNS(root, A_NS.p, "clrMapOvr");
	if (!clrMapOvr) return null;
	const override = firstChildNS(clrMapOvr, A_NS.a, "overrideClrMapping");
	if (!override) return null; // <a:masterClrMapping/> or empty → inherit
	const out = emptyClrMap();
	for (const attr of Array.from(override.attributes)) {
		out[attr.localName ?? attr.name] = attr.value;
	}
	return out;
}

// Resolve a <a:schemeClr val="..."/> to `#RRGGBB`, or null if it can't be
// resolved.
export function resolveSchemeClr(slot: string, clrMap: ClrMap, theme: ThemeColors): string | null {
	const resolvedSlot = MAPPABLE_SLOTS.has(slot) ? (clrMap[slot] ?? slot) : slot;
	return theme.scheme[resolvedSlot as SchemeSlot] ?? null;
}

// ECMA-376 §20.1.10.47 ST_PresetColorVal — the 140 canonical preset color
// names plus common aliases ("darkBlue" ↔ "dkBlue", "lightGray" ↔ "ltGray",
// "mediumBlue" ↔ "medBlue", "grey" ↔ "gray") that appear in PPTX in the
// wild. Values are the CSS/W3C named-color hex codes. Keys are lowercased
// for case-insensitive lookup.
const PRESET_COLORS: Record<string, string> = {
	aliceblue: '#F0F8FF', antiquewhite: '#FAEBD7', aqua: '#00FFFF', aquamarine: '#7FFFD4',
	azure: '#F0FFFF', beige: '#F5F5DC', bisque: '#FFE4C4', black: '#000000',
	blanchedalmond: '#FFEBCD', blue: '#0000FF', blueviolet: '#8A2BE2', brown: '#A52A2A',
	burlywood: '#DEB887', cadetblue: '#5F9EA0', chartreuse: '#7FFF00', chocolate: '#D2691E',
	coral: '#FF7F50', cornflowerblue: '#6495ED', cornsilk: '#FFF8DC', crimson: '#DC143C',
	cyan: '#00FFFF', darkblue: '#00008B', darkcyan: '#008B8B', darkgoldenrod: '#B8860B',
	darkgray: '#A9A9A9', darkgrey: '#A9A9A9', darkgreen: '#006400', darkkhaki: '#BDB76B',
	darkmagenta: '#8B008B', darkolivegreen: '#556B2F', darkorange: '#FF8C00', darkorchid: '#9932CC',
	darkred: '#8B0000', darksalmon: '#E9967A', darkseagreen: '#8FBC8F', darkslateblue: '#483D8B',
	darkslategray: '#2F4F4F', darkslategrey: '#2F4F4F', darkturquoise: '#00CED1',
	darkviolet: '#9400D3', deeppink: '#FF1493', deepskyblue: '#00BFFF', dimgray: '#696969',
	dimgrey: '#696969', dkblue: '#00008B', dkcyan: '#008B8B', dkgoldenrod: '#B8860B',
	dkgray: '#A9A9A9', dkgreen: '#006400', dkkhaki: '#BDB76B', dkmagenta: '#8B008B',
	dkolivegreen: '#556B2F', dkorange: '#FF8C00', dkorchid: '#9932CC', dkred: '#8B0000',
	dksalmon: '#E9967A', dkseagreen: '#8FBC8F', dkslateblue: '#483D8B', dkslategray: '#2F4F4F',
	dkslategrey: '#2F4F4F', dkturquoise: '#00CED1', dkviolet: '#9400D3',
	dodgerblue: '#1E90FF', firebrick: '#B22222', floralwhite: '#FFFAF0', forestgreen: '#228B22',
	fuchsia: '#FF00FF', gainsboro: '#DCDCDC', ghostwhite: '#F8F8FF', gold: '#FFD700',
	goldenrod: '#DAA520', gray: '#808080', grey: '#808080', green: '#008000',
	greenyellow: '#ADFF2F', honeydew: '#F0FFF0', hotpink: '#FF69B4', indianred: '#CD5C5C',
	indigo: '#4B0082', ivory: '#FFFFF0', khaki: '#F0E68C', lavender: '#E6E6FA',
	lavenderblush: '#FFF0F5', lawngreen: '#7CFC00', lemonchiffon: '#FFFACD', lightblue: '#ADD8E6',
	lightcoral: '#F08080', lightcyan: '#E0FFFF', lightgoldenrodyellow: '#FAFAD2',
	lightgray: '#D3D3D3', lightgrey: '#D3D3D3', lightgreen: '#90EE90', lightpink: '#FFB6C1',
	lightsalmon: '#FFA07A', lightseagreen: '#20B2AA', lightskyblue: '#87CEFA',
	lightslategray: '#778899', lightslategrey: '#778899', lightsteelblue: '#B0C4DE',
	lightyellow: '#FFFFE0', lime: '#00FF00', limegreen: '#32CD32', linen: '#FAF0E6',
	ltblue: '#ADD8E6', ltcoral: '#F08080', ltcyan: '#E0FFFF',
	ltgoldenrodyellow: '#FAFAD2', ltgray: '#D3D3D3', ltgreen: '#90EE90', ltpink: '#FFB6C1',
	ltsalmon: '#FFA07A', ltseagreen: '#20B2AA', ltskyblue: '#87CEFA', ltslategray: '#778899',
	ltslategrey: '#778899', ltsteelblue: '#B0C4DE', ltyellow: '#FFFFE0',
	magenta: '#FF00FF', maroon: '#800000', medaquamarine: '#66CDAA', medblue: '#0000CD',
	medorchid: '#BA55D3', medpurple: '#9370DB', medseagreen: '#3CB371', medslateblue: '#7B68EE',
	medspringgreen: '#00FA9A', medturquoise: '#48D1CC', medvioletred: '#C71585',
	mediumaquamarine: '#66CDAA', mediumblue: '#0000CD', mediumorchid: '#BA55D3',
	mediumpurple: '#9370DB', mediumseagreen: '#3CB371', mediumslateblue: '#7B68EE',
	mediumspringgreen: '#00FA9A', mediumturquoise: '#48D1CC', mediumvioletred: '#C71585',
	midnightblue: '#191970', mintcream: '#F5FFFA', mistyrose: '#FFE4E1',
	moccasin: '#FFE4B5', navajowhite: '#FFDEAD', navy: '#000080', oldlace: '#FDF5E6',
	olive: '#808000', olivedrab: '#6B8E23', orange: '#FFA500', orangered: '#FF4500',
	orchid: '#DA70D6', palegoldenrod: '#EEE8AA', palegreen: '#98FB98', paleturquoise: '#AFEEEE',
	palevioletred: '#DB7093', papayawhip: '#FFEFD5', peachpuff: '#FFDAB9', peru: '#CD853F',
	pink: '#FFC0CB', plum: '#DDA0DD', powderblue: '#B0E0E6', purple: '#800080',
	red: '#FF0000', rosybrown: '#BC8F8F', royalblue: '#4169E1', saddlebrown: '#8B4513',
	salmon: '#FA8072', sandybrown: '#F4A460', seagreen: '#2E8B57', seashell: '#FFF5EE',
	sienna: '#A0522D', silver: '#C0C0C0', skyblue: '#87CEEB', slateblue: '#6A5ACD',
	slategray: '#708090', slategrey: '#708090', snow: '#FFFAFA', springgreen: '#00FF7F',
	steelblue: '#4682B4', tan: '#D2B48C', teal: '#008080', thistle: '#D8BFD8',
	tomato: '#FF6347', turquoise: '#40E0D0', violet: '#EE82EE', wheat: '#F5DEB3',
	white: '#FFFFFF', whitesmoke: '#F5F5F5', yellow: '#FFFF00', yellowgreen: '#9ACD32',
};

function presetColorHex(name: string | null): string | null {
	if (!name) return null;
	return PRESET_COLORS[name.toLowerCase()] ?? null;
}

// HSL (OOXML §20.1.2.3.14) → RGB hex.
//   hue: 1/60000ths of a degree (0..21600000, i.e. 60000 × 360)
//   sat: per-mille (0..100000)
//   lum: per-mille (0..100000)
// Internally we normalise to the 0..1 representation used by the shared
// hue2rgb helper. Sanity check: hue=0, sat=100000, lum=50000 → pure red
// (#FF0000). hue=21600000 (360°) wraps to 0 and yields the same red.
function hslOoxmlToHex(hue: number, sat: number, lum: number): string {
	const h = ((hue / 60000) % 360 + 360) % 360 / 360;
	const s = Math.max(0, Math.min(1, sat / 100000));
	const l = Math.max(0, Math.min(1, lum / 100000));
	if (s === 0) {
		const v = Math.round(l * 255);
		return `#${to2(v)}${to2(v)}${to2(v)}`;
	}
	const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
	const p = 2 * l - q;
	const r = hue2rgb(p, q, h + 1 / 3);
	const g = hue2rgb(p, q, h);
	const b = hue2rgb(p, q, h - 1 / 3);
	return `#${to2(Math.round(r * 255))}${to2(Math.round(g * 255))}${to2(Math.round(b * 255))}`;
}

function hue2rgb(p: number, q: number, t: number): number {
	if (t < 0) t += 1;
	if (t > 1) t -= 1;
	if (t < 1 / 6) return p + (q - p) * 6 * t;
	if (t < 1 / 2) return q;
	if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
	return p;
}

function to2(n: number): string {
	return n.toString(16).padStart(2, "0");
}

function parsePercent(s: string | null): number | null {
	if (s == null) return null;
	// OOXML percent values: "50000" (per-mille) or "50%".
	if (s.endsWith("%")) {
		const n = Number(s.slice(0, -1));
		return Number.isNaN(n) ? null : n / 100;
	}
	const n = Number(s);
	return Number.isNaN(n) ? null : n / 100000;
}

// Shared color resolution for any DrawingML color-bearing element
// (<a:srgbClr>, <a:schemeClr>, <a:prstClr>, <a:sysClr>, <a:scrgbClr>, <a:hslClr>).
// Returns the resolved hex color, the modifier ops for later re-application,
// the underlying scheme slot (when source was schemeClr — so downstream code
// can re-resolve after clrMap changes), and a merged alpha value.
export interface ResolvedColor {
	colorHex: string;
	schemeSlot: string | null;
	mods: ColorMods | null;
	alpha: number | null;
}

export function resolveColorElement(
	el: Element,
	clrMap: ClrMap,
	theme: ThemeColors | undefined,
): ResolvedColor | null {
	if (el.namespaceURI !== A_NS.a) return null;
	const mods = parseColorMods(el);
	const alpha = extractAlpha(el);

	switch (el.localName) {
		case 'srgbClr': {
			const v = el.getAttribute("val");
			if (!v || !/^[0-9a-fA-F]{6}$/.test(v)) return null;
			return {
				colorHex: applyMods(`#${v}`, mods),
				schemeSlot: null,
				mods,
				alpha,
			};
		}
		case 'schemeClr': {
			const slot = el.getAttribute("val");
			if (!slot) return null;
			if (!theme) {
				// Not yet resolvable — but still record the scheme slot so
				// downstream passes can finish resolution.
				return { colorHex: '', schemeSlot: slot, mods, alpha };
			}
			const base = resolveSchemeClr(slot, clrMap, theme);
			if (!base) return { colorHex: '', schemeSlot: slot, mods, alpha };
			return { colorHex: applyMods(base, mods), schemeSlot: slot, mods, alpha };
		}
		case 'prstClr': {
			const name = el.getAttribute("val");
			const hex = presetColorHex(name);
			if (!hex) return null;
			return { colorHex: applyMods(hex, mods), schemeSlot: null, mods, alpha };
		}
		case 'sysClr': {
			const base = resolveSysClr(el);
			if (!base) return null;
			return { colorHex: applyMods(base, mods), schemeSlot: null, mods, alpha };
		}
		case 'scrgbClr': {
			// ECMA-376 §20.1.2.3.30 — r/g/b are per-mille sRGB components,
			// 0..100000 (0% .. 100%). Convert via Math.round(n/100000 * 255).
			// Sanity: r=100000 → 1.0 * 255 = 255 (0xFF); r=50000 → 127.
			const r = parsePercent(el.getAttribute("r"));
			const g = parsePercent(el.getAttribute("g"));
			const b = parsePercent(el.getAttribute("b"));
			if (r == null || g == null || b == null) return null;
			const hex = `#${to2(Math.round(r * 255))}${to2(Math.round(g * 255))}${to2(Math.round(b * 255))}`;
			return { colorHex: applyMods(hex, mods), schemeSlot: null, mods, alpha };
		}
		case 'hslClr': {
			const hue = Number(el.getAttribute("hue")) || 0;
			const sat = Number(el.getAttribute("sat")) || 0;
			const lum = Number(el.getAttribute("lum")) || 0;
			const hex = hslOoxmlToHex(hue, sat, lum);
			return { colorHex: applyMods(hex, mods), schemeSlot: null, mods, alpha };
		}
	}
	return null;
}

function extractAlpha(el: Element): number | null {
	for (const c of Array.from(el.children)) {
		if (c.namespaceURI !== A_NS.a) continue;
		if (c.localName === 'alpha') {
			const v = c.getAttribute("val");
			if (v == null) continue;
			const n = Number(v);
			if (!Number.isNaN(n)) return Math.max(0, Math.min(1, n / 100000));
		}
	}
	return null;
}
