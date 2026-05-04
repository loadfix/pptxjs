import type { Paragraph, Run, TextRun } from '../presentation-parser';
import { emuToPx } from './geom';
import { wrapInHyperlink } from './hyperlink';

// Counters for auto-numbered bullets within the current shape, keyed by
// paragraph level. Reset at each new shape.
export type AutoNumState = Map<number, number>;

export function renderParagraph(p: Paragraph, autoNumState: AutoNumState, hyperlinkUrls: Map<string, string>): HTMLElement {
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
	const ls = p.style.lineSpacing;
	if (ls) {
		if (ls.kind === 'pct') {
			// pct is per-mille (e.g. 150000 = 150% = line-height: 1.5).
			el.style.lineHeight = `${ls.value / 100000}`;
		} else {
			// pts is in hundredths-of-a-point (spcPts val="2400" = 24pt).
			el.style.lineHeight = `${ls.value / 100}pt`;
		}
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
		el.appendChild(renderRun(run, hyperlinkUrls));
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

export function renderRun(run: Run, hyperlinkUrls: Map<string, string>): HTMLElement {
	// Wave 2 — render BreakRun as <br>, FieldRun via field substitution.
	if (run.kind === 'break') {
		return document.createElement("br");
	}
	if (run.kind === 'field') {
		// TODO(P10): resolve field value. For now emit the fallback text plus a
		// data-pptx-field marker so the P10 fields agent can post-process.
		// slidenum in particular needs the slide index injected by the caller.
		const el = document.createElement("span");
		el.setAttribute("data-pptx-field", run.fieldType);
		el.textContent = run.fallbackText;
		applyRunStyle(el, run.style);
		return wrapInHyperlink(el, run.style.hyperlinkRId, hyperlinkUrls);
	}
	// kind === 'text'
	const el = document.createElement("span");
	el.textContent = run.text;
	applyRunStyle(el, run.style);
	return wrapInHyperlink(el, run.style.hyperlinkRId, hyperlinkUrls);
}

function applyRunStyle(el: HTMLElement, s: TextRun['style']): void {
	if (s.bold) el.style.fontWeight = "bold";
	if (s.italic) el.style.fontStyle = "italic";
	if (s.sizeHundredths != null) el.style.fontSize = `${s.sizeHundredths / 100}pt`;

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

	// Language attribute (not a style).
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
