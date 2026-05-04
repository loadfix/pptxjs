import type { TableShape, TableCell } from '../presentation-parser';
import { emuToPx, positionStyle, transformStyle } from './geom';
import { renderParagraph, type AutoNumState } from './text';
import { fillToCssBackground } from './fill-utils';

export function renderTable(t: TableShape, cls: string): HTMLElement {
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

export function renderCell(cell: TableCell): HTMLTableCellElement {
	const td = document.createElement("td");
	td.style.border = "1px solid #ccc";
	td.style.padding = "4px";
	td.style.verticalAlign = "top";
	// Cell fills are currently typed as SolidFill only; route through
	// fillToCssBackground for consistency with shape rendering.
	const bg = fillToCssBackground(cell.fill, new Map());
	if (bg) td.style.background = bg;
	if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
	if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
	const autoNumState: AutoNumState = new Map();
	for (const p of cell.paragraphs) {
		td.appendChild(renderParagraph(p, autoNumState));
	}
	return td;
}
