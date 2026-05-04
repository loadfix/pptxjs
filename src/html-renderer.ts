import type { Options } from './pptx-preview';
import type { Presentation } from './presentation';
import type { ShapeLike, Shape, PicShape, TableShape, TableCell, Paragraph, Run } from './presentation-parser';

// English Metric Units per pixel at 96 DPI. 914400 EMU = 1 inch = 96 px.
const EMU_PER_PX = 9525;
const emuToPx = (emu: number) => emu / EMU_PER_PX;

export class HtmlRenderer {
	async render(presentation: Presentation, options: Options): Promise<Node[]> {
		const out: Node[] = [];

		const slideW = emuToPx(presentation.slideSize.cx);
		const slideH = emuToPx(presentation.slideSize.cy);
		out.push(makeStyleNode(options.className, slideW, slideH));

		for (const slide of presentation.slides) {
			const section = document.createElement("section");
			section.className = `${options.className}-slide`;
			section.dataset.slideIndex = String(slide.index);
			if (slide.background?.kind === 'solid') {
				section.style.background = slide.background.colorHex;
			}
			for (const shape of slide.shapes) {
				section.appendChild(renderShapeLike(shape, options.className));
			}
			out.push(section);
		}

		return out;
	}
}

function renderShapeLike(shape: ShapeLike, cls: string): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls);
	if (shape.kind === 'table') return renderTable(shape, cls);
	return renderShape(shape, cls);
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
	const autoNumState: AutoNumState = new Map();
	for (const p of cell.paragraphs) {
		td.appendChild(renderParagraph(p, autoNumState));
	}
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

function renderShape(shape: Shape, cls: string): HTMLElement {
	const el = document.createElement("div");
	el.className = `${cls}-shape`;
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
	const autoNumState: AutoNumState = new Map();
	for (const p of shape.paragraphs) {
		el.appendChild(renderParagraph(p, autoNumState));
	}
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

function renderParagraph(p: Paragraph, autoNumState: AutoNumState): HTMLElement {
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
.${cls}-shape, .${cls}-pic, .${cls}-table {
	box-sizing: border-box;
}
`;
	return style;
}
