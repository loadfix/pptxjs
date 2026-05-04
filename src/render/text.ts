import type { Paragraph, Run, TextRun, Slide } from '../presentation-parser';
import { emuToPx } from './geom';
import { withAlphaHex } from '../color-math';
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
): HTMLElement {
	const el = document.createElement("p");
	el.style.margin = "0";

	// Left margin — prefer parsed marL (EMU), else fall back to level * 24px.
	if (p.style.marL != null) {
		el.style.marginLeft = `${emuToPx(p.style.marL)}px`;
	} else if (p.level > 0) {
		el.style.marginLeft = `${p.level * 24}px`;
	}
	if (p.style.indent != null) {
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

// Subset of OOXML auto-number types. See ECMA-376 §21.1.2.1 for the full
// list; these cover what python-pptx / Office routinely emit.
export function formatAutoNum(type: string, n: number): string {
	switch (type) {
		case "arabicPeriod": return `${n}.`;
		case "arabicParenR": return `${n})`;
		case "arabicParenBoth": return `(${n})`;
		case "arabicPlain": return `${n}`;
		case "alphaLcPeriod": return `${toAlpha(n).toLowerCase()}.`;
		case "alphaUcPeriod": return `${toAlpha(n)}.`;
		case "alphaLcParenR": return `${toAlpha(n).toLowerCase()})`;
		case "alphaUcParenR": return `${toAlpha(n)})`;
		case "romanLcPeriod": return `${toRoman(n).toLowerCase()}.`;
		case "romanUcPeriod": return `${toRoman(n)}.`;
		default: return `${n}.`;
	}
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

	// Color + alpha. If alpha is set, fold it into an `#RRGGBBAA` value
	// (CSS 8-hex); otherwise emit the plain hex color.
	if (s.colorHex) {
		el.style.color = withAlphaHex(s.colorHex, s.alpha);
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

