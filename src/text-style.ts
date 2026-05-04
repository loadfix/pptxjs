// Text-style inheritance for PPTX. A run's effective properties come from a
// chain: run → paragraph → shape → layout placeholder → master placeholder →
// master txStyles (titleStyle/bodyStyle/otherStyle). Each level supplies a
// `<a:defRPr>` (or `<a:rPr>` for the run itself) with a subset of attributes.

import { A_NS } from './namespaces';
import { firstChildNS } from './xml-utils';
import { ColorMods } from './color-math';
import { ClrMap, ThemeColors, resolveColorElement } from './theme';

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
	// Underline — u="sng" on <a:rPr>. OOXML defines many variants; we pass
	// through the raw token (narrowed to a known set when it matches).
	underline: 'none' | 'sng' | 'dbl' | 'heavy' | 'dotted' | 'dash' | 'wavy' | null;
	// Strike-through — strike="sngStrike"/"dblStrike"/"noStrike" on <a:rPr>,
	// normalized to 'sng' | 'dbl' | 'none'.
	strike: 'none' | 'sng' | 'dbl' | null;
	// baseline shift in per-mille (30000 = +30% super, -25000 = -25% sub).
	baseline: number | null;
	// Letter spacing (spc attr) in 1/100 pt.
	letterSpacingHundredths: number | null;
	// Kerning minimum (kern attr) in 1/100 pt.
	kernHundredths: number | null;
	// Alpha channel 0..1 merged across color sources (<a:alpha val="N"/> is per-mille).
	alpha: number | null;
	// Hyperlink relationship id from `<a:hlinkClick r:id="...">`.
	hyperlinkRId: string | null;
	// Language code from `<a:rPr lang="en-US">`.
	lang: string | null;
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
		underline: null,
		strike: null,
		baseline: null,
		letterSpacingHundredths: null,
		kernHundredths: null,
		alpha: null,
		hyperlinkRId: null,
		lang: null,
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
		underline: over.underline ?? under.underline,
		strike: over.strike ?? under.strike,
		baseline: over.baseline ?? under.baseline,
		letterSpacingHundredths: over.letterSpacingHundredths ?? under.letterSpacingHundredths,
		kernHundredths: over.kernHundredths ?? under.kernHundredths,
		// Alpha follows the color source when the color was overridden; otherwise
		// take over's alpha if present.
		alpha: overHasColor ? (over.alpha ?? null) : (over.alpha ?? under.alpha),
		hyperlinkRId: over.hyperlinkRId ?? under.hyperlinkRId,
		lang: over.lang ?? under.lang,
	};
}

const KNOWN_UNDERLINES = new Set([
	'none', 'sng', 'dbl', 'heavy', 'dotted', 'dotted', 'dash', 'wavy',
	// OOXML dash-variants grouped into 'dash' for rendering simplicity
	'dashLong', 'dashLongHeavy', 'dashHeavy', 'dotDash', 'dotDotDash',
	'dottedHeavy', 'wavyHeavy', 'wavyDbl',
]);

function normalizeUnderline(u: string): RunStyle['underline'] {
	if (u === 'none') return 'none';
	if (u === 'sng') return 'sng';
	if (u === 'dbl' || u === 'wavyDbl') return 'dbl';
	if (u === 'heavy' || u === 'dashHeavy' || u === 'dottedHeavy' || u === 'wavyHeavy') return 'heavy';
	if (u === 'dotted' || u === 'dotDash' || u === 'dotDotDash') return 'dotted';
	if (u === 'dash' || u === 'dashLong' || u === 'dashLongHeavy') return 'dash';
	if (u === 'wavy') return 'wavy';
	if (KNOWN_UNDERLINES.has(u)) return 'sng';
	return null;
}

function normalizeStrike(s: string): RunStyle['strike'] {
	if (s === 'noStrike') return 'none';
	if (s === 'sngStrike') return 'sng';
	if (s === 'dblStrike') return 'dbl';
	return null;
}

// Parse a `<a:rPr>` or `<a:defRPr>` element. All fields are optional. If
// `clrMap` / `theme` are provided, concrete color elements are resolved via
// `resolveColorElement`; otherwise only srgbClr/schemeClr are parsed and
// left for later resolution.
export function parseRunProps(
	el: Element | null,
	clrMap?: ClrMap,
	theme?: ThemeColors,
): RunStyle {
	const s = emptyRunStyle();
	if (!el) return s;
	const sz = el.getAttribute("sz");
	if (sz) s.sizeHundredths = Number(sz);
	const b = el.getAttribute("b");
	if (b != null) s.bold = b === "1";
	const i = el.getAttribute("i");
	if (i != null) s.italic = i === "1";

	const u = el.getAttribute("u");
	if (u) s.underline = normalizeUnderline(u);

	const strike = el.getAttribute("strike");
	if (strike) s.strike = normalizeStrike(strike);

	const baseline = el.getAttribute("baseline");
	if (baseline != null && baseline !== "") {
		const n = Number(baseline);
		if (!Number.isNaN(n)) s.baseline = n;
	}

	const spc = el.getAttribute("spc");
	if (spc != null && spc !== "") {
		const n = Number(spc);
		if (!Number.isNaN(n)) s.letterSpacingHundredths = n;
	}

	const kern = el.getAttribute("kern");
	if (kern != null && kern !== "") {
		const n = Number(kern);
		if (!Number.isNaN(n)) s.kernHundredths = n;
	}

	const lang = el.getAttribute("lang");
	if (lang) s.lang = lang;

	const solidFill = firstChildNS(el, A_NS.a, "solidFill");
	if (solidFill) {
		// Pick the first color-bearing child of the solidFill.
		const colorChild = firstColorChild(solidFill);
		if (colorChild) {
			const r = resolveColorElement(colorChild, clrMap ?? {}, theme);
			if (r) {
				s.colorHex = r.colorHex;
				s.colorSchemeSlot = r.schemeSlot;
				s.colorMods = r.mods;
				s.alpha = r.alpha;
			}
		}
	}

	// Hyperlink on run: <a:rPr><a:hlinkClick r:id="rIdN"/></a:rPr>.
	const hlinkClick = firstChildNS(el, A_NS.a, "hlinkClick");
	if (hlinkClick) {
		const rid = hlinkClick.getAttributeNS(A_NS.r, "id");
		if (rid) s.hyperlinkRId = rid;
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

// Return the first child of `el` that looks like a color-bearing DrawingML
// element. Supports srgbClr, schemeClr, prstClr, sysClr, scrgbClr, hslClr.
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
	// Line spacing and space-before/after can be expressed as per-mille percent
	// or as hundredths-of-a-point absolute.
	lineSpacing: LineSpacing | null;
	spaceBefore: LineSpacing | null;
	spaceAfter: LineSpacing | null;
	// Right-to-left paragraph flag (<a:pPr rtl="1">).
	rtl: boolean | null;
	// Tab stops (<a:tabLst><a:tab pos="..." algn="..."/>). `pos` is EMU.
	tabs: TabStop[] | null;
}

export type LineSpacing = { kind: 'pct'; value: number } | { kind: 'pts'; value: number };

export interface TabStop {
	pos: number;
	align: 'l' | 'ctr' | 'r' | 'dec' | 'clear';
}

export interface BulletMarker {
	// Bullet text size as a per-mille percent of the run's size.
	sizePct: number | null;
	colorHex: string | null;
	fontFamily: string | null;
}

export type BulletDef =
	| { kind: "none" }
	| ({ kind: "char"; char: string } & BulletMarker)
	| ({ kind: "autoNum"; type: string; startAt: number | null } & BulletMarker)
	// Image bullet — <a:buBlip><a:blip r:embed="rId"/></a:buBlip>. `rId`
	// resolves through the slide/master/layout embedUrls map at render time.
	| ({ kind: "blip"; rId: string } & BulletMarker);

export function emptyParaStyle(): ParaStyle {
	return {
		align: null,
		marL: null,
		indent: null,
		bullet: null,
		lineSpacing: null,
		spaceBefore: null,
		spaceAfter: null,
		rtl: null,
		tabs: null,
	};
}

export function mergeParaStyle(under: ParaStyle, over: ParaStyle): ParaStyle {
	return {
		align: over.align ?? under.align,
		marL: over.marL ?? under.marL,
		indent: over.indent ?? under.indent,
		bullet: over.bullet ?? under.bullet,
		lineSpacing: over.lineSpacing ?? under.lineSpacing,
		spaceBefore: over.spaceBefore ?? under.spaceBefore,
		spaceAfter: over.spaceAfter ?? under.spaceAfter,
		rtl: over.rtl ?? under.rtl,
		tabs: over.tabs ?? under.tabs,
	};
}

// Parse <a:lnSpc>|<a:spcBef>|<a:spcAft> wrapper element.
function parseSpcWrapper(wrap: Element | null): LineSpacing | null {
	if (!wrap) return null;
	const pct = firstChildNS(wrap, A_NS.a, "spcPct");
	if (pct) {
		const v = pct.getAttribute("val");
		if (v != null) {
			const n = Number(v);
			if (!Number.isNaN(n)) return { kind: 'pct', value: n };
		}
	}
	const pts = firstChildNS(wrap, A_NS.a, "spcPts");
	if (pts) {
		const v = pts.getAttribute("val");
		if (v != null) {
			const n = Number(v);
			if (!Number.isNaN(n)) return { kind: 'pts', value: n };
		}
	}
	return null;
}

function parseTabLst(lst: Element | null): TabStop[] | null {
	if (!lst) return null;
	const out: TabStop[] = [];
	for (const t of Array.from(lst.children)) {
		if (t.namespaceURI !== A_NS.a || t.localName !== 'tab') continue;
		const pos = Number(t.getAttribute("pos")) || 0;
		const algnRaw = t.getAttribute("algn") || 'l';
		const algn: TabStop['align'] =
			algnRaw === 'ctr' || algnRaw === 'r' || algnRaw === 'dec' || algnRaw === 'clear' ? algnRaw : 'l';
		out.push({ pos, align: algn });
	}
	return out.length ? out : null;
}

// Parse `<a:lvl1pPr>` etc. (or an inline `<a:pPr>`) for paragraph-level
// properties. Returns an empty style when the element is null.
export function parseParaProps(
	el: Element | null,
	clrMap?: ClrMap,
	theme?: ThemeColors,
): ParaStyle {
	const s = emptyParaStyle();
	if (!el) return s;

	const algn = el.getAttribute("algn");
	if (algn === "l" || algn === "ctr" || algn === "r" || algn === "just") s.align = algn;

	const marL = el.getAttribute("marL");
	if (marL != null) s.marL = Number(marL);
	const indent = el.getAttribute("indent");
	if (indent != null) s.indent = Number(indent);

	const rtl = el.getAttribute("rtl");
	if (rtl != null) s.rtl = rtl === "1";

	s.lineSpacing = parseSpcWrapper(firstChildNS(el, A_NS.a, "lnSpc"));
	s.spaceBefore = parseSpcWrapper(firstChildNS(el, A_NS.a, "spcBef"));
	s.spaceAfter = parseSpcWrapper(firstChildNS(el, A_NS.a, "spcAft"));

	s.tabs = parseTabLst(firstChildNS(el, A_NS.a, "tabLst"));

	// Bullet style decorations (size / color / font). Captured independently
	// of kind — they modify whichever marker type ends up chosen.
	let bulletSizePct: number | null = null;
	let bulletColorHex: string | null = null;
	let bulletFontFamily: string | null = null;

	const buSzPct = firstChildNS(el, A_NS.a, "buSzPct");
	if (buSzPct) {
		const v = buSzPct.getAttribute("val");
		if (v != null) {
			const n = Number(v);
			if (!Number.isNaN(n)) bulletSizePct = n;
		}
	}
	const buClr = firstChildNS(el, A_NS.a, "buClr");
	if (buClr) {
		const colorEl = firstColorChild(buClr);
		if (colorEl) {
			const r = resolveColorElement(colorEl, clrMap ?? {}, theme);
			if (r) bulletColorHex = r.colorHex;
		}
	}
	const buFont = firstChildNS(el, A_NS.a, "buFont");
	if (buFont) {
		const typeface = buFont.getAttribute("typeface");
		if (typeface) bulletFontFamily = typeface;
	}

	// Bullet markers. Only one of these is set per paragraph; check in
	// preference order: explicit none, char, autoNum, blip.
	if (firstChildNS(el, A_NS.a, "buNone")) {
		s.bullet = { kind: "none" };
	} else {
		const buChar = firstChildNS(el, A_NS.a, "buChar");
		const buAutoNum = firstChildNS(el, A_NS.a, "buAutoNum");
		const buBlip = firstChildNS(el, A_NS.a, "buBlip");
		if (buChar) {
			const c = buChar.getAttribute("char");
			if (c) {
				s.bullet = {
					kind: "char",
					char: c,
					sizePct: bulletSizePct,
					colorHex: bulletColorHex,
					fontFamily: bulletFontFamily,
				};
			}
		} else if (buAutoNum) {
			s.bullet = {
				kind: "autoNum",
				type: buAutoNum.getAttribute("type") ?? "arabicPeriod",
				startAt: buAutoNum.getAttribute("startAt") ? Number(buAutoNum.getAttribute("startAt")) : null,
				sizePct: bulletSizePct,
				colorHex: bulletColorHex,
				fontFamily: bulletFontFamily,
			};
		} else if (buBlip) {
			const blip = firstChildNS(buBlip, A_NS.a, "blip");
			const rId = blip ? blip.getAttributeNS(A_NS.r, "embed") : null;
			if (rId) {
				s.bullet = {
					kind: "blip",
					rId,
					sizePct: bulletSizePct,
					colorHex: bulletColorHex,
					fontFamily: bulletFontFamily,
				};
			}
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
export function parseLevelStyles(
	container: Element | null,
	clrMap?: ClrMap,
	theme?: ThemeColors,
): LevelStyles {
	const out = emptyLevelStyles();
	if (!container) return out;
	for (let lvl = 1; lvl <= 9; lvl++) {
		const lvlEl = firstChildNS(container, A_NS.a, `lvl${lvl}pPr`);
		if (!lvlEl) continue;
		const defRPr = firstChildNS(lvlEl, A_NS.a, "defRPr");
		out[lvl - 1] = {
			run: defRPr ? parseRunProps(defRPr, clrMap, theme) : null,
			para: parseParaProps(lvlEl, clrMap, theme),
		};
	}
	return out;
}

// -- Body properties (<a:bodyPr>) ---------------------------------------------

export interface BodyProperties {
	lInsEmu: number | null;
	tInsEmu: number | null;
	rInsEmu: number | null;
	bInsEmu: number | null;
	anchor: 't' | 'ctr' | 'b' | 'just' | 'dist' | null;
	wrap: 'none' | 'square' | null;
	vert: 'horz' | 'vert' | 'vert270' | 'wordArtVert' | 'eaVert' | 'mongolianVert' | 'wordArtVertRtl' | null;
	numCol: number | null;
	spcColEmu: number | null;
	autofit:
		| { kind: 'none' }
		| { kind: 'normAutofit'; fontScale: number | null; lnSpcReduction: number | null }
		| { kind: 'spAutoFit' }
		| null;
}

export function emptyBodyProperties(): BodyProperties {
	return {
		lInsEmu: null,
		tInsEmu: null,
		rInsEmu: null,
		bInsEmu: null,
		anchor: null,
		wrap: null,
		vert: null,
		numCol: null,
		spcColEmu: null,
		autofit: null,
	};
}

export function parseBodyPr(bodyPrEl: Element | null): BodyProperties | null {
	if (!bodyPrEl) return null;
	const out = emptyBodyProperties();

	const numAttr = (n: string): number | null => {
		const v = bodyPrEl.getAttribute(n);
		if (v == null || v === "") return null;
		const num = Number(v);
		return Number.isNaN(num) ? null : num;
	};

	out.lInsEmu = numAttr("lIns");
	out.tInsEmu = numAttr("tIns");
	out.rInsEmu = numAttr("rIns");
	out.bInsEmu = numAttr("bIns");

	const anchor = bodyPrEl.getAttribute("anchor");
	if (anchor === 't' || anchor === 'ctr' || anchor === 'b' || anchor === 'just' || anchor === 'dist') {
		out.anchor = anchor;
	}
	const wrap = bodyPrEl.getAttribute("wrap");
	if (wrap === 'none' || wrap === 'square') out.wrap = wrap;

	const vert = bodyPrEl.getAttribute("vert");
	if (
		vert === 'horz' || vert === 'vert' || vert === 'vert270'
		|| vert === 'wordArtVert' || vert === 'eaVert'
		|| vert === 'mongolianVert' || vert === 'wordArtVertRtl'
	) out.vert = vert;

	out.numCol = numAttr("numCol");
	out.spcColEmu = numAttr("spcCol");

	// Autofit children — first match wins.
	const normAutofit = firstChildNS(bodyPrEl, A_NS.a, "normAutofit");
	if (normAutofit) {
		const fs = normAutofit.getAttribute("fontScale");
		const ls = normAutofit.getAttribute("lnSpcReduction");
		out.autofit = {
			kind: 'normAutofit',
			fontScale: fs != null && fs !== "" ? Number(fs) : null,
			lnSpcReduction: ls != null && ls !== "" ? Number(ls) : null,
		};
	} else if (firstChildNS(bodyPrEl, A_NS.a, "spAutoFit")) {
		out.autofit = { kind: 'spAutoFit' };
	} else if (firstChildNS(bodyPrEl, A_NS.a, "noAutofit")) {
		out.autofit = { kind: 'none' };
	}

	return out;
}
