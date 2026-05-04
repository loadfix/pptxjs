import type { Paragraph, Run, TextRun } from '../presentation-parser';

// Counters for auto-numbered bullets within the current shape, keyed by
// paragraph level. Reset at each new shape.
export type AutoNumState = Map<number, number>;

export function renderParagraph(p: Paragraph, autoNumState: AutoNumState): HTMLElement {
	const el = document.createElement("p");
	el.style.margin = "0";
	if (p.level > 0) el.style.marginLeft = `${p.level * 24}px`;

	const align = p.style.align;
	if (align === "ctr") el.style.textAlign = "center";
	else if (align === "r") el.style.textAlign = "right";
	else if (align === "just") el.style.textAlign = "justify";

	const bullet = p.style.bullet;
	if (bullet && bullet.kind !== "none") {
		const marker = document.createElement("span");
		marker.style.marginRight = "0.4em";
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
		el.appendChild(renderRun(run));
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

export function renderRun(run: Run): HTMLElement {
	// Wave 2 — render BreakRun as <br>, FieldRun via field substitution.
	if (run.kind === 'break') {
		return document.createElement("br");
	}
	if (run.kind === 'field') {
		const el = document.createElement("span");
		el.textContent = run.fallbackText;
		applyRunStyle(el, run.style);
		return el;
	}
	// kind === 'text'
	const el = document.createElement("span");
	el.textContent = run.text;
	applyRunStyle(el, run.style);
	return el;
}

function applyRunStyle(el: HTMLElement, s: TextRun['style']): void {
	if (s.bold) el.style.fontWeight = "bold";
	if (s.italic) el.style.fontStyle = "italic";
	if (s.sizeHundredths != null) el.style.fontSize = `${s.sizeHundredths / 100}pt`;
	if (s.colorHex) el.style.color = s.colorHex;
	if (s.fontFamily) el.style.fontFamily = s.fontFamily;
	// Wave 2 P11 — expose <a:rPr lang=> as `lang` so screen readers can pick
	// the right pronunciation for the run's text.
	if (s.lang) el.setAttribute("lang", s.lang);
}
