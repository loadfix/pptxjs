import type { TableShape, TableCell } from '../presentation-parser';
import { emuToPx, positionStyle } from './geom';
import { renderParagraph, type AutoNumState, type FieldContext } from './text';

export function renderTable(t: TableShape, cls: string, fieldCtx?: FieldContext): HTMLElement {
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
			tr.appendChild(renderCell(cell, fieldCtx));
		}
		table.appendChild(tr);
	}

	wrap.appendChild(table);
	return wrap;
}

export function renderCell(cell: TableCell, fieldCtx?: FieldContext): HTMLTableCellElement {
	const td = document.createElement("td");
	td.style.border = "1px solid #ccc";
	td.style.padding = "4px";
	td.style.verticalAlign = "top";
	if (cell.fill?.kind === 'solid') td.style.background = cell.fill.colorHex;
	if (cell.gridSpan > 1) td.colSpan = cell.gridSpan;
	if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
	const autoNumState: AutoNumState = new Map();
	for (const p of cell.paragraphs) {
		td.appendChild(renderParagraph(p, autoNumState, fieldCtx));
	}
	return td;
}
