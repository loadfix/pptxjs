import type { TableShape, TableCell, TableCellBorders } from '../presentation-parser';
import type { TableStyle, TableBandStyle, TableBorders } from '../table-style';
import type { LineStyle, Fill } from '../fill';
import { emuToPx, positionStyle, transformStyle } from './geom';
import { renderParagraph, type AutoNumState, type FieldContext } from './text';
import { solidColorFromFill } from './fill-utils';

// DrawingML <a:prstDash> values → CSS border-style. "solid" is the default
// when no dash is specified on the line.
function dashToCss(dash: LineStyle['dash']): string {
	switch (dash) {
		case 'dash':
		case 'lgDash':
		case 'sysDash':
			return 'dashed';
		case 'dot':
		case 'sysDot':
			return 'dotted';
		case 'dashDot':
		case 'lgDashDot':
		case 'sysDashDot':
			return 'dashed';
		case 'solid':
		case null:
		case undefined:
		default:
			return 'solid';
	}
}

// Build a CSS `border` shorthand from a LineStyle. Returns null when the
// line resolves to a no-op (noFill or unresolvable colour).
function borderCss(line: LineStyle | null): string | null {
	if (!line) return null;
	// A NoFill <a:ln> collapses to widthEmu=0 at parse time — signal "no border".
	if (line.widthEmu === 0) return 'none';
	const color = solidColorFromFill(line.fill ?? null);
	if (!color) return null;
	const widthPx = line.widthEmu != null ? Math.max(emuToPx(line.widthEmu), 0.5) : 1;
	return `${widthPx}px ${dashToCss(line.dash)} ${color}`;
}

// SVG stroke-dasharray derived from a LineStyle. Units are multiples of
// stroke width (matching what PowerPoint effectively does). Returns null
// for solid lines so the attribute can be omitted.
function dashToSvgArray(dash: LineStyle['dash']): string | null {
	switch (dash) {
		case 'dash':
		case 'sysDash':
			return '4 3';
		case 'lgDash':
			return '8 3';
		case 'dot':
		case 'sysDot':
			return '1 3';
		case 'dashDot':
		case 'sysDashDot':
			return '4 3 1 3';
		case 'lgDashDot':
			return '8 3 1 3';
		default:
			return null;
	}
}

// Build an <svg> overlay drawing a diagonal line inside the cell. The
// viewBox is a fixed 100x100 and preserveAspectRatio="none" stretches
// the line with whatever size the cell takes. Returns null when the line
// resolves to a no-op.
function diagonalSvg(line: LineStyle, corners: 'tlbr' | 'blTr'): SVGSVGElement | null {
	if (line.widthEmu === 0) return null;
	const color = solidColorFromFill(line.fill ?? null);
	if (!color) return null;
	const widthPx = line.widthEmu != null ? Math.max(emuToPx(line.widthEmu), 0.5) : 1;

	const svgNs = 'http://www.w3.org/2000/svg';
	const svg = document.createElementNS(svgNs, 'svg');
	svg.setAttribute('viewBox', '0 0 100 100');
	svg.setAttribute('preserveAspectRatio', 'none');
	svg.setAttribute('width', '100%');
	svg.setAttribute('height', '100%');
	svg.style.position = 'absolute';
	svg.style.inset = '0';
	svg.style.pointerEvents = 'none';

	const lineEl = document.createElementNS(svgNs, 'line');
	if (corners === 'tlbr') {
		lineEl.setAttribute('x1', '0');
		lineEl.setAttribute('y1', '0');
		lineEl.setAttribute('x2', '100');
		lineEl.setAttribute('y2', '100');
	} else {
		lineEl.setAttribute('x1', '0');
		lineEl.setAttribute('y1', '100');
		lineEl.setAttribute('x2', '100');
		lineEl.setAttribute('y2', '0');
	}
	lineEl.setAttribute('stroke', color);
	lineEl.setAttribute('stroke-width', String(widthPx));
	lineEl.setAttribute('vector-effect', 'non-scaling-stroke');
	const dashArr = dashToSvgArray(line.dash);
	if (dashArr) lineEl.setAttribute('stroke-dasharray', dashArr);
	svg.appendChild(lineEl);
	return svg;
}

// Merge band b into a. Later arguments win (used for the cascade).
function mergeBand(a: TableBandStyle, b: TableBandStyle): TableBandStyle {
	return {
		borders: {
			l: b.borders.l ?? a.borders.l,
			r: b.borders.r ?? a.borders.r,
			t: b.borders.t ?? a.borders.t,
			b: b.borders.b ?? a.borders.b,
			insideH: b.borders.insideH ?? a.borders.insideH,
			insideV: b.borders.insideV ?? a.borders.insideV,
			tlbr: b.borders.tlbr ?? a.borders.tlbr,
			blTr: b.borders.blTr ?? a.borders.blTr,
		},
		fill: b.fill ?? a.fill,
		textColorHex: b.textColorHex ?? a.textColorHex,
	};
}

function emptyBandInline(): TableBandStyle {
	return {
		borders: { l: null, r: null, t: null, b: null, insideH: null, insideV: null, tlbr: null, blTr: null },
		fill: null,
		textColorHex: null,
	};
}

// Derive the effective band style for a specific cell. Cascades:
//   wholeTbl → row banding (band1H/band2H for inner rows; firstRow/lastRow
//   override the banding on the edges) → column banding (band1V/band2V
//   for inner columns; firstCol/lastCol override on edges) → corner cells
//   (nwCell/neCell/swCell/seCell) when both surrounding row+col flags fire.
function resolveBandForCell(
	style: TableStyle,
	flags: TableShape['tableFlags'],
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
): TableBandStyle {
	let out = emptyBandInline();
	out = mergeBand(out, style.wholeTbl);

	// Row banding. When firstRow is on and this is row 0, skip the band
	// colouring for that row and let firstRow win instead. Similarly for
	// lastRow.
	const isFirstRow = flags.firstRow && rowIndex === 0;
	const isLastRow = flags.lastRow && rowIndex === rowCount - 1;
	if (flags.bandRow && !isFirstRow && !isLastRow) {
		// PowerPoint: band1H covers rows 1,3,5,... when firstRow is present
		// (index 0 is the header), or 0,2,4,... when it isn't.
		const offset = flags.firstRow ? 1 : 0;
		const banded = (rowIndex - offset);
		if (banded >= 0) {
			out = mergeBand(out, banded % 2 === 0 ? style.band1H : style.band2H);
		}
	}
	if (isFirstRow) out = mergeBand(out, style.firstRow);
	if (isLastRow) out = mergeBand(out, style.lastRow);

	// Column banding, mirroring the row-banding logic. Applied before
	// firstCol/lastCol so those still win on the edges.
	const isFirstCol = flags.firstCol && colIndex === 0;
	const isLastCol = flags.lastCol && colIndex === colCount - 1;
	if (flags.bandCol && !isFirstCol && !isLastCol) {
		const offset = flags.firstCol ? 1 : 0;
		const banded = (colIndex - offset);
		if (banded >= 0) {
			out = mergeBand(out, banded % 2 === 0 ? style.band1V : style.band2V);
		}
	}
	if (isFirstCol) out = mergeBand(out, style.firstCol);
	if (isLastCol) out = mergeBand(out, style.lastCol);

	// Corner cells — apply last so they override the row/col bands at the
	// four table corners. Both the row-edge flag and the column-edge flag
	// must be set for the corner style to fire.
	if (isFirstRow && isFirstCol) out = mergeBand(out, style.nwCell);
	if (isFirstRow && isLastCol) out = mergeBand(out, style.neCell);
	if (isLastRow && isFirstCol) out = mergeBand(out, style.swCell);
	if (isLastRow && isLastCol) out = mergeBand(out, style.seCell);

	return out;
}

// Convert an EMU inset 4-tuple to a CSS padding shorthand.
function insetsToPadding(i: { l: number; r: number; t: number; b: number } | null): string {
	if (!i) {
		// PowerPoint defaults: 0.1" left/right, 0.05" top/bottom.
		return `${emuToPx(45720)}px ${emuToPx(91440)}px`;
	}
	return `${emuToPx(i.t)}px ${emuToPx(i.r)}px ${emuToPx(i.b)}px ${emuToPx(i.l)}px`;
}

function anchorToVAlign(a: 't' | 'ctr' | 'b' | null): string {
	if (a === 'ctr') return 'middle';
	if (a === 'b') return 'bottom';
	return 'top';
}

function fillToBackground(fill: Fill | null): string | null {
	if (!fill) return null;
	if (fill.kind === 'solid') return fill.colorHex;
	if (fill.kind === 'none') return 'transparent';
	// Gradient/pattern/blip fills in cells — TODO, render as no background.
	return null;
}

// Determine the border-on-a-side value for a cell, given:
//   - the cell's own per-side border (wins when present),
//   - the band's per-side border for the outer edge,
//   - the band's insideH/insideV for interior edges.
function resolveSideBorder(
	cellSide: LineStyle | null,
	bandOuterSide: LineStyle | null,
	bandInsideSide: LineStyle | null,
	isOuter: boolean,
): LineStyle | null {
	if (cellSide) return cellSide;
	if (isOuter) return bandOuterSide ?? bandInsideSide;
	return bandInsideSide ?? bandOuterSide;
}

export function renderTable(
	t: TableShape,
	cls: string,
	tableStyles: Map<string, TableStyle> | null,
	hyperlinkUrls: Map<string, string>,
	embedUrls: Map<string, string>,
	fieldCtx?: FieldContext,
): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-table`;
	Object.assign(wrap.style, positionStyle(t.x, t.y, t.cx, t.cy));
	Object.assign(wrap.style, transformStyle(t.rotation60k, t.flipH, t.flipV));

	const table = document.createElement("table");
	Object.assign(table.style, {
		width: "100%",
		height: "100%",
		borderCollapse: "collapse",
		tableLayout: "fixed",
	});

	// Accessibility: explicit role="table" (redundant for the native element
	// but harmless, and guards against CSS display overrides). aria-label
	// prefers the title; falls back to the shape's alt/name text.
	table.setAttribute("role", "table");
	const tableLabel = t.title || t.alt || t.name;
	if (tableLabel) table.setAttribute("aria-label", tableLabel);

	if (t.colWidthsEmu.length > 0) {
		const colgroup = document.createElement("colgroup");
		for (const w of t.colWidthsEmu) {
			const col = document.createElement("col");
			col.style.width = `${emuToPx(w)}px`;
			colgroup.appendChild(col);
		}
		table.appendChild(colgroup);
	}

	// Look up the style. styleId might point at an entry that isn't in
	// the map (default theme styles like {5C22544A-...} usually aren't
	// instantiated in tableStyles.xml) — fall back to null (no cascade).
	const style = t.styleId && tableStyles ? tableStyles.get(t.styleId) ?? null : null;

	const rowCount = t.rows.length;
	const colCount = t.colWidthsEmu.length || (t.rows[0]?.cells.length ?? 0);

	for (let rowIndex = 0; rowIndex < t.rows.length; rowIndex++) {
		const row = t.rows[rowIndex];
		const tr = document.createElement("tr");
		if (row.heightEmu) tr.style.height = `${emuToPx(row.heightEmu)}px`;
		let colIndex = 0;
		for (const cell of row.cells) {
			if (cell.hMerge || cell.vMerge) { colIndex++; continue; }
			const band = style
				? resolveBandForCell(style, t.tableFlags, rowIndex, rowCount, colIndex, colCount)
				: null;
			tr.appendChild(renderCell(cell, band, rowIndex, rowCount, colIndex, colCount, hyperlinkUrls, embedUrls, fieldCtx));
			colIndex += cell.gridSpan || 1;
		}
		table.appendChild(tr);
	}

	wrap.appendChild(table);
	return wrap;
}

export function renderCell(
	cell: TableCell,
	band: TableBandStyle | null,
	rowIndex: number,
	rowCount: number,
	colIndex: number,
	colCount: number,
	hyperlinkUrls: Map<string, string>,
	embedUrls: Map<string, string>,
	fieldCtx?: FieldContext,
): HTMLTableCellElement {
	const td = document.createElement("td");

	// Padding — per-cell <a:tcPr marL/R/T/B> wins; else PowerPoint defaults.
	td.style.padding = insetsToPadding(cell.insetsEmu);
	// Vertical alignment — from <a:tcPr anchor="...">. Defaults to top.
	td.style.verticalAlign = anchorToVAlign(cell.anchor);

	// Background — cell fill wins; then band fill; else transparent.
	const cellBg = fillToBackground(cell.fill);
	if (cellBg) {
		td.style.background = cellBg;
	} else if (band) {
		const bandBg = fillToBackground(band.fill);
		if (bandBg) td.style.background = bandBg;
	}

	// Text colour — band's tcTxStyle applies if the cell didn't set its own.
	if (band?.textColorHex) td.style.color = band.textColorHex;

	// Per-side borders.
	const bandBorders: TableBorders | null = band ? band.borders : null;
	const cellBorders: TableCellBorders = cell.borders;
	const isOuterT = rowIndex === 0;
	const isOuterB = rowIndex === rowCount - 1;
	const isOuterL = colIndex === 0;
	const isOuterR = colIndex === colCount - 1;

	const sideT = resolveSideBorder(cellBorders.t, bandBorders?.t ?? null, bandBorders?.insideH ?? null, isOuterT);
	const sideB = resolveSideBorder(cellBorders.b, bandBorders?.b ?? null, bandBorders?.insideH ?? null, isOuterB);
	const sideL = resolveSideBorder(cellBorders.l, bandBorders?.l ?? null, bandBorders?.insideV ?? null, isOuterL);
	const sideR = resolveSideBorder(cellBorders.r, bandBorders?.r ?? null, bandBorders?.insideV ?? null, isOuterR);

	const bT = borderCss(sideT);
	const bB = borderCss(sideB);
	const bL = borderCss(sideL);
	const bR = borderCss(sideR);
	if (bT) td.style.borderTop = bT;
	if (bB) td.style.borderBottom = bB;
	if (bL) td.style.borderLeft = bL;
	if (bR) td.style.borderRight = bR;

	if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
	if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;

	// Diagonal borders. <a:lnTlToBr> and <a:lnBlToTr> don't map onto any
	// native <td> border property, so we overlay them as absolutely-
	// positioned SVG inside the cell. pointer-events: none lets clicks
	// still reach the cell content.
	//
	// Merged cells (Wave 6 A4): when a diagonal is declared on a cell
	// with gridSpan>1 or rowSpan>1, this approach Just Works — setting
	// colSpan/rowSpan on the <td> stretches the cell's own box across
	// the full merged rect, and the SVG's viewBox="0 0 100 100" +
	// preserveAspectRatio="none" scales the line to fill that box. The
	// (0,0)→(100,100) / (0,100)→(100,0) endpoints therefore land on the
	// corners of the merged rectangle, not the original single cell.
	// Continuation cells (hMerge/vMerge) are skipped in renderTable
	// above so only the spanning cell renders a diagonal.
	const tlbrLine = cellBorders.tlbr ?? bandBorders?.tlbr ?? null;
	const blTrLine = cellBorders.blTr ?? bandBorders?.blTr ?? null;
	if (tlbrLine || blTrLine) td.style.position = 'relative';
	if (tlbrLine) {
		const svg = diagonalSvg(tlbrLine, 'tlbr');
		if (svg) td.appendChild(svg);
	}
	if (blTrLine) {
		const svg = diagonalSvg(blTrLine, 'blTr');
		if (svg) td.appendChild(svg);
	}

	const autoNumState: AutoNumState = new Map();
	for (const p of cell.paragraphs) {
		td.appendChild(renderParagraph(p, autoNumState, hyperlinkUrls, fieldCtx, undefined, embedUrls));
	}
	return td;
}
