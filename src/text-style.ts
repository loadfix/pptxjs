// Text-style inheritance for PPTX. A run's effective properties come from a
// chain: run → paragraph → shape → layout placeholder → master placeholder →
// master txStyles (titleStyle/bodyStyle/otherStyle). Each level supplies a
// `<a:defRPr>` (or `<a:rPr>` for the run itself) with a subset of attributes.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ColorMods, parseColorMods, applyMods } from './color-math';

export interface RunStyle {
	sizeHundredths: number | null;  // `sz` attribute, hundredths of a point
	bold: boolean | null;
	italic: boolean | null;
	colorHex: string | null;        // `#RRGGBB` — resolved from srgbClr *or* schemeClr.
	// If the source was a <a:schemeClr val="..."/>, the slot name is recorded
	// here so later passes can apply theme/clrMap resolution. Once resolved
	// into `colorHex`, this is cleared.
	colorSchemeSlot: string | null;
	colorMods: ColorMods | null;    // tint/shade/lumMod/lumOff captured with the schemeClr
	fontFamily: string | null;
}

export function emptyRunStyle(): RunStyle {
	return {
		sizeHundredths: null,
		bold: null,
		italic: null,
		colorHex: null,
		colorSchemeSlot: null,
		colorMods: null,
		fontFamily: null,
	};
}

// Merge `over` on top of `under` — non-null fields in `over` win. Color is
// merged as a unit: if `over` specifies either colorHex or colorSchemeSlot,
// both fields from `over` replace both from `under`, so a schemeClr override
// doesn't leave a stale srgbClr behind.
export function mergeRunStyle(under: RunStyle, over: RunStyle): RunStyle {
	const overHasColor = over.colorHex != null || over.colorSchemeSlot != null;
	return {
		sizeHundredths: over.sizeHundredths ?? under.sizeHundredths,
		bold: over.bold ?? under.bold,
		italic: over.italic ?? under.italic,
		colorHex: overHasColor ? over.colorHex : under.colorHex,
		colorSchemeSlot: overHasColor ? over.colorSchemeSlot : under.colorSchemeSlot,
		colorMods: overHasColor ? over.colorMods : under.colorMods,
		fontFamily: over.fontFamily ?? under.fontFamily,
	};
}

// Parse a `<a:rPr>` or `<a:defRPr>` element. All fields are optional.
export function parseRunProps(el: Element | null): RunStyle {
	const s = emptyRunStyle();
	if (!el) return s;
	const sz = el.getAttribute("sz");
	if (sz) s.sizeHundredths = Number(sz);
	const b = el.getAttribute("b");
	if (b != null) s.bold = b === "1";
	const i = el.getAttribute("i");
	if (i != null) s.italic = i === "1";

	const solidFill = firstChildNS(el, A_NS.a, "solidFill");
	if (solidFill) {
		const srgb = firstChildNS(solidFill, A_NS.a, "srgbClr");
		if (srgb) {
			const v = srgb.getAttribute("val");
			if (v && /^[0-9a-fA-F]{6}$/.test(v)) {
				// Apply any modifier children directly on the srgb element.
				s.colorHex = applyMods(`#${v}`, parseColorMods(srgb));
			}
		} else {
			const scheme = firstChildNS(solidFill, A_NS.a, "schemeClr");
			if (scheme) {
				const v = scheme.getAttribute("val");
				if (v) {
					s.colorSchemeSlot = v;
					s.colorMods = parseColorMods(scheme);
				}
			}
		}
	}

	const latin = firstChildNS(el, A_NS.a, "latin");
	if (latin) {
		// Keep theme references ("+mj-lt", "+mn-lt") as-is; resolve them
		// later against the presentation's theme fontScheme.
		const t = latin.getAttribute("typeface");
		if (t) s.fontFamily = t;
	}
	return s;
}

// Paragraph-level properties. Inherits through the same chain as RunStyle.
export interface ParaStyle {
	// `algn` values in OOXML: "l" | "ctr" | "r" | "just" | "dist" | "thaiDist" | "justLow"
	// Rendered as `text-align: left|center|right|justify`.
	align: "l" | "ctr" | "r" | "just" | null;
	// Left margin and first-line indent in EMU.
	marL: number | null;
	indent: number | null;
	// Bullet: 'none' (explicit <a:buNone/>) | { char } | { autoNum }. null means
	// "not set" — inherit. 'none' means "explicitly suppressed".
	bullet: BulletDef | null;
}

export type BulletDef =
	| { kind: "none" }
	| { kind: "char"; char: string }
	| { kind: "autoNum"; type: string; startAt: number | null };

export function emptyParaStyle(): ParaStyle {
	return { align: null, marL: null, indent: null, bullet: null };
}

export function mergeParaStyle(under: ParaStyle, over: ParaStyle): ParaStyle {
	return {
		align: over.align ?? under.align,
		marL: over.marL ?? under.marL,
		indent: over.indent ?? under.indent,
		bullet: over.bullet ?? under.bullet,
	};
}

// Parse `<a:lvl1pPr>` etc. (or an inline `<a:pPr>`) for paragraph-level
// properties. Returns an empty style when the element is null.
export function parseParaProps(el: Element | null): ParaStyle {
	const s = emptyParaStyle();
	if (!el) return s;

	const algn = el.getAttribute("algn");
	if (algn === "l" || algn === "ctr" || algn === "r" || algn === "just") s.align = algn;

	const marL = el.getAttribute("marL");
	if (marL != null) s.marL = Number(marL);
	const indent = el.getAttribute("indent");
	if (indent != null) s.indent = Number(indent);

	// Bullet markers. Only one of these is set per paragraph; check in
	// preference order: explicit none, char, autoNum.
	if (firstChildNS(el, A_NS.a, "buNone")) {
		s.bullet = { kind: "none" };
	} else {
		const buChar = firstChildNS(el, A_NS.a, "buChar");
		const buAutoNum = firstChildNS(el, A_NS.a, "buAutoNum");
		if (buChar) {
			const c = buChar.getAttribute("char");
			if (c) s.bullet = { kind: "char", char: c };
		} else if (buAutoNum) {
			s.bullet = {
				kind: "autoNum",
				type: buAutoNum.getAttribute("type") ?? "arabicPeriod",
				startAt: buAutoNum.getAttribute("startAt") ? Number(buAutoNum.getAttribute("startAt")) : null,
			};
		}
	}
	return s;
}

export interface LevelStyle {
	run: RunStyle | null;
	para: ParaStyle | null;
}

// Indexed by level 0..8. Each entry combines the level's paragraph- and
// run-level defaults from `<a:lvl1pPr>` (attributes + child `<a:defRPr>`).
export type LevelStyles = (LevelStyle | null)[];

export function emptyLevelStyles(): LevelStyles {
	return [null, null, null, null, null, null, null, null, null];
}

// Parse an <a:lstStyle> or <p:titleStyle>/<p:bodyStyle>/<p:otherStyle> — any
// element whose children are <a:lvl1pPr>..<a:lvl9pPr> (and optionally
// <a:defPPr>).
export function parseLevelStyles(container: Element | null): LevelStyles {
	const out = emptyLevelStyles();
	if (!container) return out;
	for (let lvl = 1; lvl <= 9; lvl++) {
		const lvlEl = firstChildNS(container, A_NS.a, `lvl${lvl}pPr`);
		if (!lvlEl) continue;
		const defRPr = firstChildNS(lvlEl, A_NS.a, "defRPr");
		out[lvl - 1] = {
			run: defRPr ? parseRunProps(defRPr) : null,
			para: parseParaProps(lvlEl),
		};
	}
	return out;
}
