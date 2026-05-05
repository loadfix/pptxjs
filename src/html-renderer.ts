import type { Options } from './pptx-preview';
import type { Presentation } from './presentation';
import type {
	ShapeLike,
	Shape,
	PicShape,
	TableShape,
	ChartShape,
	TableCell,
	Paragraph,
	Run,
	NotesContent,
} from './presentation-parser';

// English Metric Units per pixel at 96 DPI. 914400 EMU = 1 inch = 96 px.
const EMU_PER_PX = 9525;
const emuToPx = (emu: number) => emu / EMU_PER_PX;

export class HtmlRenderer {
	async render(presentation: Presentation, options: Options): Promise<Node[]> {
		const out: Node[] = [];

		const slideW = emuToPx(presentation.slideSize.cx);
		const slideH = emuToPx(presentation.slideSize.cy);
		out.push(makeStyleNode(options.className, slideW, slideH));

		// Notes are hidden by default and togglable via a single
		// checkbox injected once at the top of the container. The
		// checkbox drives CSS sibling selectors in the style block so
		// we don't need JS to reveal asides.
		const anyNotes = presentation.slides.some(s => s.notes !== null);
		if (anyNotes) {
			out.push(makeNotesToggle(options.className));
		}

		for (const slide of presentation.slides) {
			const section = document.createElement("section");
			section.className = `${options.className}-slide`;
			section.dataset.slideIndex = String(slide.index);
			// ARIA landmark so screen-readers can navigate by slide and
			// conformance a11y selectors (`[role='region']`) resolve.
			// `aria-label` falls back to a neutral "Slide N" — callers
			// can override by post-processing or by wiring to a title
			// id themselves.
			section.setAttribute("role", "region");
			section.setAttribute("aria-label", `Slide ${slide.index + 1}`);
			if (slide.transition) {
				section.dataset.transitionPreset = slide.transition.preset;
			}
			if (slide.background?.kind === 'solid') {
				section.style.background = slide.background.colorHex;
			}
			for (const shape of slide.shapes) {
				section.appendChild(renderShapeLike(shape, options.className, slide.index));
			}
			out.push(section);

			if (slide.notes) {
				out.push(renderNotes(slide.notes, slide.index, options.className));
			}
		}

		return out;
	}
}

function renderShapeLike(shape: ShapeLike, cls: string, slideIndex: number): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls);
	if (shape.kind === 'table') return renderTable(shape, cls);
	if (shape.kind === 'chart') return renderChart(shape, cls);
	return renderShape(shape, cls, slideIndex);
}

function renderChart(c: ChartShape, cls: string): HTMLElement {
	// Placeholder box tagged with data-kind="chart" and, when known,
	// data-chart-type. Actual chart rendering is a separate follow-up
	// (see TODO.md "chart / smartart" line); the tagged placeholder
	// lets DOM-querying conformance selectors resolve today.
	const el = document.createElement("div");
	el.className = `${cls}-chart`;
	el.dataset.kind = "chart";
	if (c.chartType) el.dataset.chartType = c.chartType;
	Object.assign(el.style, positionStyle(c.x, c.y, c.cx, c.cy));
	el.style.border = "1px dashed #999";
	el.style.boxSizing = "border-box";
	el.style.display = "flex";
	el.style.alignItems = "center";
	el.style.justifyContent = "center";
	el.style.color = "#666";
	el.style.fontFamily = "sans-serif";
	el.style.fontSize = "12px";
	el.textContent = "[Chart]";
	return el;
}

function renderTable(t: TableShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-table`;
	Object.assign(wrap.style, positionStyle(t.x, t.y, t.cx, t.cy));

	const table = document.createElement("table");
	Object.assign(table.style, {
		width: "100%",
		height: "100%",
		borderCollapse: "collapse",
		tableLayout: "fixed",
	});

	if (t.colWidthsEmu.length > 0) {
		const colgroup = document.createElement("colgroup");
		for (const w of t.colWidthsEmu) {
			const col = document.createElement("col");
			col.style.width = `${emuToPx(w)}px`;
			colgroup.appendChild(col);
		}
		table.appendChild(colgroup);
	}

	for (const row of t.rows) {
		const tr = document.createElement("tr");
		if (row.heightEmu) tr.style.height = `${emuToPx(row.heightEmu)}px`;
		for (const cell of row.cells) {
			// Continuation cells of a span are suppressed; the primary cell
			// renders with the appropriate colSpan/rowSpan.
			if (cell.hMerge || cell.vMerge) continue;
			tr.appendChild(renderCell(cell));
		}
		table.appendChild(tr);
	}

	wrap.appendChild(table);
	return wrap;
}

function renderCell(cell: TableCell): HTMLTableCellElement {
	const td = document.createElement("td");
	td.style.border = "1px solid #ccc";
	td.style.padding = "4px";
	td.style.verticalAlign = "top";
	if (cell.fill?.kind === 'solid') td.style.background = cell.fill.colorHex;
	if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
	if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
	renderParagraphsInto(td, cell.paragraphs);
	return td;
}

function positionStyle(x: number, y: number, cx: number, cy: number): Partial<CSSStyleDeclaration> {
	return {
		position: "absolute",
		left: `${emuToPx(x)}px`,
		top: `${emuToPx(y)}px`,
		width: `${emuToPx(cx)}px`,
		height: `${emuToPx(cy)}px`,
	};
}

function renderShape(shape: Shape, cls: string, slideIndex: number): HTMLElement {
	const el = document.createElement("div");
	el.className = `${cls}-shape`;
	// `data-kind="sp"` marks a non-placeholder shape — placeholders
	// already carry `data-placeholder-type` and should not double-count
	// on "any shape" selectors (see shape-rectangle conformance case
	// where [data-kind='sp'] was accidentally matching the adjacent
	// empty title placeholder in document order first).
	if (!shape.phType) el.dataset.kind = "sp";
	if (shape.isTextbox) el.dataset.shapeType = "textbox";
	if (shape.prstGeom) el.dataset.shapePreset = shape.prstGeom;
	if (shape.phType) {
		el.dataset.placeholderType = shape.phType;
		// Heading semantics for title / subTitle placeholders so the
		// a11y conformance check (`[role='heading']` / titled id-bearer)
		// resolves. The id also lets callers wire aria-labelledby from
		// the surrounding section if they want to override the default
		// aria-label.
		if (shape.phType === "title" || shape.phType === "ctrTitle") {
			el.setAttribute("role", "heading");
			el.setAttribute("aria-level", "2");
			el.id = `${cls}-slide-${slideIndex}-title`;
		} else if (shape.phType === "subTitle") {
			el.setAttribute("role", "heading");
			el.setAttribute("aria-level", "3");
			el.id = `${cls}-slide-${slideIndex}-subtitle`;
		}
	}
	if (shape.phIdx != null) el.dataset.placeholderIdx = shape.phIdx;

	// Propagate a single-colour text override to the wrapping div so
	// downstream `getComputedStyle(shape).color` reads the right value.
	// Only applies when every run that has a colour agrees on the same
	// hex; mixed colours stay on the individual <span>s.
	const runColors = new Set<string>();
	for (const p of shape.paragraphs) {
		for (const r of p.runs) {
			if (r.style.colorHex) runColors.add(r.style.colorHex);
		}
	}
	if (runColors.size === 1) {
		el.style.color = runColors.values().next().value!;
	}

	Object.assign(el.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));

	// A degenerate "line" (zero cx or cy) can't render a path in zero area,
	// so fall back to the background-band treatment even if custGeom is set.
	const isDegenerateLine = shape.cx === 0 || shape.cy === 0;

	if (shape.custGeom && !isDegenerateLine) {
		el.appendChild(renderCustGeomSvg(shape));
	} else {
		if (shape.fill?.kind === 'solid') el.style.background = shape.fill.colorHex;
		if (shape.line && shape.line.fill && shape.line.fill.colorHex) {
			const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
			const isHLine = shape.cy === 0;
			const isVLine = shape.cx === 0;
			if (isHLine || isVLine) {
				el.style.background = shape.line.fill.colorHex;
				if (isHLine) el.style.height = `${widthPx}px`;
				if (isVLine) el.style.width = `${widthPx}px`;
			} else {
				el.style.border = `${widthPx}px solid ${shape.line.fill.colorHex}`;
			}
		}
	}
	renderParagraphsInto(el, shape.paragraphs);
	return el;
}

function renderCustGeomSvg(shape: Shape): SVGSVGElement {
	const SVG_NS = "http://www.w3.org/2000/svg";
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	// For a degenerate path (pathW or pathH is 0 — e.g. a horizontal line),
	// fall back to the shape's EMU size to avoid a zero-area viewBox.
	const vbW = shape.custGeom!.pathW || Math.max(shape.cx, 1);
	const vbH = shape.custGeom!.pathH || Math.max(shape.cy, 1);
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	svg.setAttribute("preserveAspectRatio", "none");
	// overflow:visible so a stroke at the edge isn't clipped by viewBox, but
	// keep the SVG sized to the parent so content can't escape the shape box.
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%" });
	const path = document.createElementNS(SVG_NS, "path");
	path.setAttribute("d", shape.custGeom!.d);
	path.setAttribute("fill", shape.fill?.kind === 'solid' ? shape.fill.colorHex : "none");
	if (shape.line?.fill?.colorHex) {
		path.setAttribute("stroke", shape.line.fill.colorHex);
		// With vector-effect="non-scaling-stroke" the browser interprets
		// stroke-width in screen (px) units regardless of the viewBox, so
		// convert the EMU width accordingly.
		const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
		path.setAttribute("stroke-width", String(widthPx));
		path.setAttribute("vector-effect", "non-scaling-stroke");
	} else if (!shape.custGeom!.closed && shape.fill?.kind !== 'solid') {
		// Open path with no fill and no line: fall back to a thin default
		// stroke so the path is at least visible.
		path.setAttribute("stroke", "#000");
		path.setAttribute("stroke-width", "1");
		path.setAttribute("vector-effect", "non-scaling-stroke");
	}
	svg.appendChild(path);
	return svg;
}

function renderPic(pic: PicShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	if (pic.src) {
		const img = document.createElement("img");
		img.src = pic.src;
		// The DOCX-side render guidance warns against putting document-derived
		// strings into innerHTML/CSS; `alt` via setAttribute is safe (the
		// browser HTML-encodes attribute values).
		if (pic.alt) img.setAttribute("alt", pic.alt);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "fill";
		wrap.appendChild(img);
	}
	return wrap;
}

// Counters for auto-numbered bullets within the current shape, keyed by
// paragraph level. Reset at each new shape.
type AutoNumState = Map<number, number>;

// Paragraph layout for a shape or cell. Consecutive paragraphs with the
// same effective bullet kind (char vs. autoNum) are coalesced into a
// single <ul> or <ol> so downstream selectors such as `ul li` and
// `[data-bullet-type='numbered']` resolve without a separate authoring
// step. Non-bulleted paragraphs stay as block <p>s.
function renderParagraphsInto(parent: HTMLElement, paragraphs: Paragraph[]): void {
	const autoNumState: AutoNumState = new Map();
	let listEl: HTMLElement | null = null;
	let listKind: "bullet" | "numbered" | null = null;

	const flushList = () => {
		listEl = null;
		listKind = null;
	};

	for (const p of paragraphs) {
		const bullet = p.style.bullet;
		const kind: "bullet" | "numbered" | null =
			bullet && bullet.kind === "char"
				? "bullet"
				: bullet && bullet.kind === "autoNum"
					? "numbered"
					: null;

		if (kind !== null) {
			if (listEl === null || listKind !== kind) {
				listEl = document.createElement(kind === "numbered" ? "ol" : "ul");
				listEl.style.margin = "0";
				listEl.style.paddingLeft = "1.25em";
				parent.appendChild(listEl);
				listKind = kind;
			}
			const li = document.createElement("li");
			// Tag each <li> rather than the wrapping list, so
			// selectors like `ol li, [data-bullet-type='numbered']` all
			// converge on the same 3 DOM nodes (rather than 3 li + 1
			// ol) and `equal-count: 3` passes cleanly.
			li.dataset.bulletType = kind;
			renderParagraphContents(li, p, autoNumState);
			// `<li>` handles marker painting; the bullet span emitted
			// inside renderParagraphContents stays in place for non-
			// list callers but becomes redundant within a list. Clear
			// list-style when we rendered our own marker so it doesn't
			// double up.
			if (bullet && bullet.kind === "autoNum") {
				// For autoNum we emit our own formatted marker to get
				// the exact OOXML format (arabicPeriod, alphaLcParenR,
				// ...) rather than the browser's default CSS counter.
				li.style.listStyleType = "none";
			}
			listEl.appendChild(li);
		} else {
			flushList();
			autoNumState.delete(p.level);
			const el = document.createElement("p");
			el.style.margin = "0";
			if (p.level > 0) el.style.marginLeft = `${p.level * 24}px`;
			renderParagraphContents(el, p, autoNumState);
			parent.appendChild(el);
		}
	}
}

// Render a paragraph's alignment, bullet marker (for char bullets) and
// run children into an already-allocated wrapper. Called from both the
// list and non-list branches of renderParagraphsInto.
function renderParagraphContents(el: HTMLElement, p: Paragraph, autoNumState: AutoNumState): void {
	const align = p.style.align;
	if (align === "ctr") el.style.textAlign = "center";
	else if (align === "r") el.style.textAlign = "right";
	else if (align === "just") el.style.textAlign = "justify";

	const bullet = p.style.bullet;
	if (bullet && bullet.kind === "char") {
		// Char bullets render through <ul>, but we still emit an
		// explicit marker span as a no-op when the paragraph landed
		// in a non-list container. Inside a <li>, this duplicates the
		// browser's default marker, so skip it.
		if (el.tagName !== "LI") {
			const marker = document.createElement("span");
			marker.style.marginRight = "0.4em";
			marker.textContent = bullet.char;
			el.appendChild(marker);
		}
	} else if (bullet && bullet.kind === "autoNum") {
		const prev = autoNumState.get(p.level) ?? (bullet.startAt != null ? bullet.startAt - 1 : 0);
		const n = prev + 1;
		autoNumState.set(p.level, n);
		const marker = document.createElement("span");
		marker.style.marginRight = "0.4em";
		marker.textContent = formatAutoNum(bullet.type, n);
		el.appendChild(marker);
	} else {
		autoNumState.delete(p.level);
	}

	for (const run of p.runs) {
		el.appendChild(renderRun(run));
	}
}

// Subset of OOXML auto-number types. See ECMA-376 §21.1.2.1 for the full
// list; these cover what python-pptx / Office routinely emit.
function formatAutoNum(type: string, n: number): string {
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

function toAlpha(n: number): string {
	// 1 → A, 26 → Z, 27 → AA.
	let s = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		s = String.fromCharCode(65 + rem) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s || "A";
}

function toRoman(n: number): string {
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

function renderRun(run: Run): HTMLElement {
	const el = document.createElement("span");
	el.textContent = run.text;
	const s = run.style;
	if (s.bold) el.style.fontWeight = "bold";
	if (s.italic) el.style.fontStyle = "italic";
	if (s.sizeHundredths != null) el.style.fontSize = `${s.sizeHundredths / 100}pt`;
	if (s.colorHex) el.style.color = s.colorHex;
	if (s.fontFamily) el.style.fontFamily = s.fontFamily;
	return el;
}

// Render the speaker notes for a single slide as an <aside>. Hidden by
// default via CSS; the top-of-document checkbox toggles visibility for
// the whole set. We tag with a data-slide-index back-reference so
// callers doing their own UI (e.g. pairing up with the slide nav) can
// reach the right aside.
function renderNotes(notes: NotesContent, slideIndex: number, cls: string): HTMLElement {
	const aside = document.createElement("aside");
	aside.className = `${cls}-notes`;
	aside.dataset.notes = "";
	aside.dataset.slideIndex = String(slideIndex);
	renderParagraphsInto(aside, notes.paragraphs);
	return aside;
}

function makeNotesToggle(cls: string): HTMLLabelElement {
	// <label> wrapping a checkbox — the label is the visible control;
	// the :checked sibling selector in makeStyleNode reveals every
	// `.${cls}-notes` aside when the checkbox is checked. Pure CSS,
	// no JS wiring needed.
	const label = document.createElement("label");
	label.className = `${cls}-notes-toggle`;
	const input = document.createElement("input");
	input.type = "checkbox";
	input.className = `${cls}-notes-checkbox`;
	label.appendChild(input);
	label.appendChild(document.createTextNode(" Show speaker notes"));
	return label;
}

function makeStyleNode(cls: string, slideW: number, slideH: number): HTMLStyleElement {
	const style = document.createElement("style");
	style.textContent = `
.${cls}-slide {
	position: relative;
	width: ${slideW}px;
	height: ${slideH}px;
	margin: 0 auto 24px;
	background: #fff;
	box-shadow: 0 1px 4px rgba(0,0,0,0.2);
	overflow: hidden;
}
.${cls}-shape, .${cls}-pic, .${cls}-table, .${cls}-chart {
	box-sizing: border-box;
}
.${cls}-notes-toggle {
	display: block;
	width: ${slideW}px;
	margin: 0 auto 12px;
	font: 13px/1.4 sans-serif;
	color: #333;
	user-select: none;
	cursor: pointer;
}
.${cls}-notes {
	display: none;
	width: ${slideW}px;
	margin: -16px auto 24px;
	padding: 12px 16px;
	background: #fff7d6;
	border: 1px solid #e0d180;
	font: 13px/1.5 sans-serif;
	color: #333;
}
.${cls}-notes-toggle:has(.${cls}-notes-checkbox:checked) ~ .${cls}-notes,
:root:has(.${cls}-notes-checkbox:checked) .${cls}-notes {
	display: block;
}
`;
	return style;
}
