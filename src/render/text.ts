import type { Paragraph, Run, TextRun, Slide } from '../presentation-parser';
import { emuToPx } from './geom';
import { wrapInHyperlink } from './hyperlink';

// Counters for auto-numbered bullets within the current shape, keyed by
// paragraph level. Reset at each new shape.
export type AutoNumState = Map<number, number>;

// Resolution context for field substitution (slidenum, datetime, ftr, hdr).
// The renderer threads one of these through paragraph/run rendering so runs
// of kind `field` can be replaced with the current slide's values.
export interface FieldContext {
	slide: Slide;
	// Presentation-level 1-based starting slide number from
	// <p:presentation firstSlideNum="N">; combines with slide.index.
	firstSlideNum: number;
}

// Body-level modifiers threaded from the enclosing shape's <a:bodyPr> through
// paragraph / run rendering. Both values come from <a:normAutofit>:
//   fontScale      — per-mille; 75000 → 0.75 multiplier on every run size.
//   lnSpcReduction — per-mille; 20000 → subtract 20% from every paragraph's
//                    effective line-height.
export interface BodyTextContext {
	fontScale: number | null;
	lnSpcReduction: number | null;
}

export function renderParagraph(
	p: Paragraph,
	autoNumState: AutoNumState,
	hyperlinkUrls: Map<string, string>,
	fieldCtx?: FieldContext,
	bodyCtx?: BodyTextContext,
	embedUrls?: Map<string, string>,
): HTMLElement {
	const el = document.createElement("p");
	el.style.margin = "0";

	// Left padding — effective left margin comes from `marL` when set, else
	// falls back to the PowerPoint default of 342900 EMU per indent level
	// (roughly 0.375"). Emit as padding-left so text-indent on the first line
	// pulls the bullet marker back into the gutter without clipping.
	const DEFAULT_MARL_PER_LEVEL_EMU = 342900;
	const effectiveMarL = p.style.marL != null
		? p.style.marL
		: (p.level > 0 ? p.level * DEFAULT_MARL_PER_LEVEL_EMU : 0);
	if (effectiveMarL > 0) {
		el.style.paddingLeft = `${emuToPx(effectiveMarL)}px`;
	}
	// First-line indent. OOXML stores `indent` as a signed EMU; negative
	// values pull the bullet marker left of the text block (hanging indent).
	if (p.style.indent != null && p.style.indent !== 0) {
		el.style.textIndent = `${emuToPx(p.style.indent)}px`;
	}

	const align = p.style.align;
	if (align === "ctr") el.style.textAlign = "center";
	else if (align === "r") el.style.textAlign = "right";
	else if (align === "just") el.style.textAlign = "justify";

	// Line spacing — percent maps to unitless line-height; points stays in pt.
	// normAutofit lnSpcReduction (per-mille) trims a percentage off whichever
	// line-height was chosen. When no explicit lineSpacing is set but a
	// reduction applies, treat the baseline as 100% so the reduction still
	// takes effect.
	const ls = p.style.lineSpacing;
	const lnReduction = bodyCtx?.lnSpcReduction;
	if (ls) {
		if (ls.kind === 'pct') {
			// pct is per-mille (e.g. 150000 = 150% = line-height: 1.5).
			const base = ls.value / 100000;
			const trimmed = lnReduction != null ? base * (1 - lnReduction / 100000) : base;
			el.style.lineHeight = `${trimmed}`;
		} else {
			// pts is in hundredths-of-a-point (spcPts val="2400" = 24pt).
			const basePt = ls.value / 100;
			const trimmedPt = lnReduction != null ? basePt * (1 - lnReduction / 100000) : basePt;
			el.style.lineHeight = `${trimmedPt}pt`;
		}
	} else if (lnReduction != null) {
		el.style.lineHeight = `${1 - lnReduction / 100000}`;
	}
	const sb = p.style.spaceBefore;
	if (sb) {
		el.style.marginTop = sb.kind === 'pct' ? `${sb.value / 100000}em` : `${sb.value / 100}pt`;
	}
	const sa = p.style.spaceAfter;
	if (sa) {
		el.style.marginBottom = sa.kind === 'pct' ? `${sa.value / 100000}em` : `${sa.value / 100}pt`;
	}

	if (p.style.rtl) el.style.direction = "rtl";

	// TODO: tab stops (ParaStyle.tabs) — HTML can't easily emulate OOXML
	// tab-stop lists with per-stop alignment. Skipped for now.

	const bullet = p.style.bullet;
	if (bullet && bullet.kind !== "none") {
		const marker = document.createElement("span");
		marker.style.marginRight = "0.4em";
		// BulletMarker fields: sizePct, colorHex, fontFamily.
		if (bullet.sizePct != null) {
			// sizePct is per-mille (e.g. 75000 = 75%).
			marker.style.fontSize = `${bullet.sizePct / 1000}%`;
		}
		if (bullet.colorHex) marker.style.color = bullet.colorHex;
		if (bullet.fontFamily) marker.style.fontFamily = bullet.fontFamily;
		if (bullet.kind === "char") {
			marker.textContent = bullet.char;
		} else if (bullet.kind === "blip") {
			// Image bullet — resolve the rId through the embedUrls map.
			// Fall back silently (empty marker) when the embed is missing.
			const url = embedUrls?.get(bullet.rId);
			if (url) {
				const img = document.createElement("img");
				img.src = url;
				img.alt = "";
				// Size the glyph to roughly the text height. Using `em` keeps it
				// in sync with the surrounding run size.
				img.style.height = "1em";
				img.style.verticalAlign = "text-bottom";
				marker.appendChild(img);
			}
		} else {
			const prev = autoNumState.get(p.level) ?? (bullet.startAt != null ? bullet.startAt - 1 : 0);
			const n = prev + 1;
			autoNumState.set(p.level, n);
			marker.textContent = formatAutoNum(bullet.type, n);
		}
		el.appendChild(marker);
	} else {
		// Reset downstream counters when a paragraph at this level is non-auto.
		autoNumState.delete(p.level);
	}

	for (const run of p.runs) {
		el.appendChild(renderRun(run, hyperlinkUrls, fieldCtx, bodyCtx));
	}
	return el;
}

// OOXML auto-number types. See ECMA-376 §21.1.2.1 (ST_TextAutonumberScheme)
// for the full list. Unhandled exotic variants fall through to `${n}.`.
export function formatAutoNum(type: string, n: number): string {
	switch (type) {
		// --- Arabic ---
		case "arabicPeriod": return `${n}.`;
		case "arabicParenR": return `${n})`;
		case "arabicParenBoth": return `(${n})`;
		case "arabicPlain": return `${n}`;
		case "arabic1Minus": return `${n}-`;
		case "arabic2Minus": return `${n}.-`;
		// Double-byte arabic — render as plain ASCII digits for now.
		case "arabicDbPeriod": return `${n}.`;
		case "arabicDbPlain": return `${n}`;

		// --- Latin alpha ---
		case "alphaLcPeriod": return `${toAlpha(n).toLowerCase()}.`;
		case "alphaUcPeriod": return `${toAlpha(n)}.`;
		case "alphaLcParenR": return `${toAlpha(n).toLowerCase()})`;
		case "alphaUcParenR": return `${toAlpha(n)})`;

		// --- Roman ---
		case "romanLcPeriod": return `${toRoman(n).toLowerCase()}.`;
		case "romanUcPeriod": return `${toRoman(n)}.`;

		// --- Circled numerals (Unicode U+2460..U+2473 for 1..20) ---
		case "circleNumDbPlain":
		case "circleNumWdBlackPlain":
		case "circleNumWdWhitePlain":
			return toCircled(n);

		// --- East-Asian numerals. Chs/Cht/Jpn share CJK digits 一,二...十. ---
		case "ea1ChsPeriod":
		case "ea1ChtPeriod":
		case "ea1JpnChsDbPeriod":
			return `${toCjkNum(n)}.`;
		case "ea1JpnKorPlain":
			return toHangulNum(n);

		// --- Hebrew letters (1..22) ---
		case "hebrew2Minus": return `${toHebrew(n)}-`;

		// --- Thai ---
		case "thaiNumPeriod": return `${toThaiDigits(n)}.`;
		case "thaiNumParenR": return `${toThaiDigits(n)})`;
		case "thaiNumParenBoth": return `(${toThaiDigits(n)})`;
		case "thaiAlphaPeriod": return `${toThaiLetter(n)}.`;
		case "thaiAlphaParenR": return `${toThaiLetter(n)})`;
		case "thaiAlphaParenBoth": return `(${toThaiLetter(n)})`;

		default: return `${n}.`;
	}
}

// 1..20 → ①..⑳ (U+2460..U+2473). Fall back to `${n}.` past 20.
export function toCircled(n: number): string {
	if (n >= 1 && n <= 20) return String.fromCharCode(0x245f + n);
	return `${n}.`;
}

// CJK numerals: 一,二,三,四,五,六,七,八,九,十. 11..19 as 十一..十九, 20 as
// 二十, 21..29 as 二十一..二十九... up to 99. Past that, fall back.
export function toCjkNum(n: number): string {
	const digits = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'];
	if (n < 1) return `${n}`;
	if (n <= 10) return n === 10 ? '十' : digits[n];
	if (n < 20) return '十' + digits[n - 10];
	if (n < 100) {
		const tens = Math.floor(n / 10);
		const ones = n % 10;
		return digits[tens] + '十' + (ones ? digits[ones] : '');
	}
	return `${n}`;
}

// Korean Hangul numerals 1..10 (하나 둘 셋 넷 다섯 여섯 일곱 여덟 아홉 열).
// Past 10, fall back to the digit.
export function toHangulNum(n: number): string {
	const table = ['하나', '둘', '셋', '넷', '다섯', '여섯', '일곱', '여덟', '아홉', '열'];
	if (n >= 1 && n <= 10) return table[n - 1];
	return `${n}`;
}

// Hebrew letters א..ת (22 letters). Fall back to `${n}.` past 22.
export function toHebrew(n: number): string {
	// 0x05D0 = א.
	if (n >= 1 && n <= 22) return String.fromCharCode(0x05cf + n);
	return `${n}`;
}

// Thai digits ๐..๙ (U+0E50..U+0E59).
export function toThaiDigits(n: number): string {
	return String(n)
		.split('')
		.map(c => /[0-9]/.test(c) ? String.fromCharCode(0x0e50 + Number(c)) : c)
		.join('');
}

// Thai consonants ก..ฮ (U+0E01..U+0E2E, 44 letters). Fall back past 44.
export function toThaiLetter(n: number): string {
	if (n >= 1 && n <= 44) return String.fromCharCode(0x0e00 + n);
	return `${n}`;
}

export function toAlpha(n: number): string {
	// 1 → A, 26 → Z, 27 → AA.
	let s = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		s = String.fromCharCode(65 + rem) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s || "A";
}

export function toRoman(n: number): string {
	const table: [number, string][] = [
		[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
		[100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
		[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"],
	];
	let out = "";
	for (const [v, s] of table) {
		while (n >= v) { out += s; n -= v; }
	}
	return out || "I";
}

export function renderRun(
	run: Run,
	hyperlinkUrls: Map<string, string>,
	fieldCtx?: FieldContext,
	bodyCtx?: BodyTextContext,
): HTMLElement {
	// Wave 2 — render BreakRun as <br>, FieldRun via field substitution.
	if (run.kind === 'break') {
		return document.createElement("br");
	}
	if (run.kind === 'field') {
		// TODO(P10): resolve field value. For now emit the fallback text plus a
		// data-pptx-field marker so the P10 fields agent can post-process.
		// slidenum in particular needs the slide index injected by the caller.
		const el = document.createElement("span");
		const resolved = resolveField(run.fieldType, run.fallbackText, fieldCtx);
		el.textContent = resolved;
		el.setAttribute("data-pptx-field", run.fieldType);
		applyRunStyle(el, run.style, bodyCtx);
		return wrapInHyperlink(el, run.style.hyperlinkRId, hyperlinkUrls);
	}
	// kind === 'text'
	const el = document.createElement("span");
	el.textContent = run.text;
	applyRunStyle(el, run.style, bodyCtx);
	return wrapInHyperlink(el, run.style.hyperlinkRId, hyperlinkUrls);
}

// Resolve a FieldRun's fieldType to actual display text. Falls through to
// `fallbackText` for unknown/unsupported variants so the slide still shows
// *something* — PowerPoint usually authors a reasonable literal there.
export function resolveField(fieldType: string, fallbackText: string, ctx?: FieldContext): string {
	if (!ctx) return fallbackText;
	const t = fieldType;
	if (t === 'slidenum' || t === 'sldNum') {
		// slide.index is 0-based within the kept (non-hidden) list; combine
		// with firstSlideNum (1-based) so index=0 → firstSlideNum.
		return String(ctx.slide.index + ctx.firstSlideNum);
	}
	if (t === 'ftr') {
		if (!ctx.slide.hf.ftr) return "";
		return ctx.slide.footerText ?? fallbackText;
	}
	if (t === 'hdr') {
		return ctx.slide.headerText ?? fallbackText;
	}
	if (t === 'datetime' || t === 'datetimeFigureOut' || /^datetime(\d+)?$/.test(t)) {
		return formatDatetimeField(t, fallbackText);
	}
	return fallbackText;
}

// Minimal mapping from OOXML `datetimeN` variants to locale-friendly output.
// We render using the browser's current locale; an authoritative format-code
// translator would be far larger than the prompt warrants. Unrecognised
// variants fall through to fallbackText.
export function formatDatetimeField(variant: string, fallbackText: string): string {
	const now = new Date();
	switch (variant) {
		case 'datetime':
		case 'datetime1':
			// "M/D/YYYY" style — the default short-date locale rendering.
			return now.toLocaleDateString();
		case 'datetime2':
			// Long date, e.g. "Monday, May 02, 2026".
			return now.toLocaleDateString(undefined, {
				weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',
			});
		case 'datetime3':
			// Abbreviated e.g. "02 May 2026".
			return now.toLocaleDateString(undefined, {
				day: '2-digit', month: 'short', year: 'numeric',
			});
		case 'datetime12':
		case 'datetime13':
			// Time-of-day variants.
			return now.toLocaleTimeString();
		case 'datetimeFigureOut':
			return now.toLocaleString();
		default:
			return fallbackText;
	}
}

function applyRunStyle(el: HTMLElement, s: TextRun['style'], bodyCtx?: BodyTextContext): void {
	if (s.bold) el.style.fontWeight = "bold";
	if (s.italic) el.style.fontStyle = "italic";
	if (s.sizeHundredths != null) {
		// normAutofit fontScale (per-mille) shrinks every run uniformly so the
		// text fits its frame. 75000 → 0.75× base size.
		const scale = bodyCtx?.fontScale != null ? bodyCtx.fontScale / 100000 : 1;
		el.style.fontSize = `${(s.sizeHundredths / 100) * scale}pt`;
	}

	// Color + alpha. If alpha is set, fold it into an rgba() value; otherwise
	// emit the plain hex color.
	if (s.colorHex) {
		if (s.alpha != null && s.alpha < 1) {
			el.style.color = hexToRgba(s.colorHex, s.alpha);
		} else {
			el.style.color = s.colorHex;
		}
	} else if (s.alpha != null && s.alpha < 1) {
		// No explicit color but alpha present — use opacity as a fallback.
		el.style.opacity = `${s.alpha}`;
	}

	if (s.fontFamily) el.style.fontFamily = s.fontFamily;

	// Text-decoration: combine underline + strikethrough.
	const decoParts: string[] = [];
	let decoStyle: string | null = null;
	if (s.underline && s.underline !== 'none') {
		decoParts.push("underline");
		switch (s.underline) {
			case 'dbl': decoStyle = "double"; break;
			case 'dotted': decoStyle = "dotted"; break;
			case 'dash': decoStyle = "dashed"; break;
			case 'wavy': decoStyle = "wavy"; break;
			// 'sng' and 'heavy' fall through to the default solid style.
		}
	}
	if (s.strike && s.strike !== 'none') {
		decoParts.push("line-through");
	}
	if (decoParts.length > 0) {
		el.style.textDecoration = decoParts.join(" ") + (decoStyle ? ` ${decoStyle}` : "");
	} else if (s.underline === 'none' && s.strike === 'none') {
		el.style.textDecoration = "none";
	} else if (s.underline === 'none' || s.strike === 'none') {
		// Explicit none on just one of the two — still suppress decoration.
		el.style.textDecoration = "none";
	}

	// Baseline shift — per-mille. Positive = super, negative = sub. Super/sub
	// text typically renders at ~58% of base size in Office; for non-standard
	// shifts we honour the absolute percentage.
	if (s.baseline != null && s.baseline !== 0) {
		el.style.verticalAlign = s.baseline > 0 ? "super" : "sub";
		const pct = Math.abs(s.baseline) / 1000; // e.g. 30000 → 30
		// If the shift is close to Office's default (~30% for super, -25% for
		// sub), use the conventional 58% scaling; otherwise use the raw value.
		const scale = pct >= 20 && pct <= 40 ? 58 : pct;
		el.style.fontSize = `${scale}%`;
	}

	// Letter spacing — spc is in 1/100 pt.
	if (s.letterSpacingHundredths != null) {
		el.style.letterSpacing = `${s.letterSpacingHundredths / 100}pt`;
	}

	// Kerning — HTML doesn't expose granular kern thresholds; when the source
	// asks for any kerning (kern attr set), enable font-kerning.
	if (s.kernHundredths != null) {
		el.style.fontKerning = "normal";
	}

	// Language attribute (not a style). <a:rPr lang=...> is exposed on the
	// rendered span so screen readers can pick the right pronunciation.
	if (s.lang) {
		el.setAttribute("lang", s.lang);
	}
}

// `#RRGGBB` + alpha in [0,1] → `rgba(r, g, b, a)`.
function hexToRgba(hex: string, alpha: number): string {
	const h = hex.startsWith("#") ? hex.slice(1) : hex;
	if (h.length !== 6) return hex;
	const r = parseInt(h.slice(0, 2), 16);
	const g = parseInt(h.slice(2, 4), 16);
	const b = parseInt(h.slice(4, 6), 16);
	return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}
