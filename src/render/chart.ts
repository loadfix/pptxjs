// Inline SVG renderer for the chart families pptxjs understands natively:
// bar / column (clustered, stacked, percentStacked), line, pie, doughnut.
// Output is a self-contained <svg> sized to the graphic frame; axes, ticks
// and a legend are drawn from the ChartModel directly. Styling is kept
// minimal — we're not trying to reproduce PowerPoint's pixel-for-pixel
// rendering, just produce something recognisable at a glance.

import type { ChartModel, ChartSeries, ChartLegendPos } from '../chart-parser';
import { SVG_NS } from './geom';

const PALETTE = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'];

function svg(el: string, attrs: Record<string, string | number>): SVGElement {
	const node = document.createElementNS(SVG_NS, el);
	for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
	return node;
}

function text(x: number, y: number, s: string, extra: Record<string, string | number> = {}): SVGElement {
	const t = svg('text', {
		x, y,
		'font-size': 10,
		'font-family': 'sans-serif',
		fill: '#333',
		...extra,
	});
	t.textContent = s;
	return t;
}

function seriesColor(s: ChartSeries, idx: number): string {
	return s.colorHex ?? PALETTE[idx % PALETTE.length];
}

// Budget the chart rectangle into margins + plot area. Extra room is
// reserved on the legend side when the legend sits beside the plot.
interface Layout {
	ml: number; mr: number; mt: number; mb: number;
	plotX: number; plotY: number; plotW: number; plotH: number;
}

function layout(model: ChartModel, width: number, height: number): Layout {
	let ml = 50, mr = 20, mt = model.title ? 60 : 30, mb = 40;
	if (model.legend === 'r' || model.legend === 'tr') mr = Math.max(mr, 100);
	if (model.legend === 'l') ml = Math.max(ml, 100);
	if (model.legend === 'b') mb = Math.max(mb, 60);
	if (model.legend === 't') mt = Math.max(mt, model.title ? 80 : 50);
	return {
		ml, mr, mt, mb,
		plotX: ml,
		plotY: mt,
		plotW: Math.max(10, width - ml - mr),
		plotH: Math.max(10, height - mt - mb),
	};
}

// Pick a "nice" tick count close to 5 for the y axis.
function niceTicks(min: number, max: number, count = 5): number[] {
	if (min === max) { min -= 1; max += 1; }
	const span = max - min;
	const step0 = span / count;
	const mag = Math.pow(10, Math.floor(Math.log10(step0)));
	const norm = step0 / mag;
	let step: number;
	if (norm < 1.5) step = 1 * mag;
	else if (norm < 3) step = 2 * mag;
	else if (norm < 7) step = 5 * mag;
	else step = 10 * mag;
	const tMin = Math.floor(min / step) * step;
	const tMax = Math.ceil(max / step) * step;
	const out: number[] = [];
	for (let v = tMin; v <= tMax + step * 0.5; v += step) out.push(Number(v.toFixed(10)));
	return out;
}

// For column/bar: compute y-domain based on grouping. Stacked sums each
// category, percentStacked is always 0..100, standard/clustered spans
// min(0, minVal)..maxVal.
function valueDomain(model: ChartModel): { min: number; max: number } {
	if (model.grouping === 'percentStacked') return { min: 0, max: 100 };
	const n = model.categories.length || (model.series[0]?.values.length ?? 0);
	if (model.grouping === 'stacked') {
		let mx = 0, mn = 0;
		for (let i = 0; i < n; i++) {
			let posSum = 0, negSum = 0;
			for (const s of model.series) {
				const v = s.values[i] ?? 0;
				if (v >= 0) posSum += v; else negSum += v;
			}
			if (posSum > mx) mx = posSum;
			if (negSum < mn) mn = negSum;
		}
		return { min: mn, max: mx || 1 };
	}
	let mx = 0, mn = 0;
	for (const s of model.series) for (const v of s.values) {
		if (v > mx) mx = v;
		if (v < mn) mn = v;
	}
	return { min: mn, max: mx || 1 };
}

function truncate(s: string, maxChars: number): string {
	if (s.length <= maxChars) return s;
	if (maxChars <= 1) return '…';
	return s.slice(0, maxChars - 1) + '…';
}

function drawLegend(root: SVGElement, model: ChartModel, L: Layout, width: number, height: number): void {
	if (!model.legend) return;
	const items = model.series.map((s, i) => ({ name: s.name || `Series ${i + 1}`, color: seriesColor(s, i) }));
	// Resolve alternate positions to r/b.
	let pos: ChartLegendPos = model.legend;
	if (pos === 'tr') pos = 'r';
	if (pos === 't') pos = 'b';

	if (pos === 'r' || pos === 'l') {
		const x = pos === 'r' ? width - L.mr + 8 : 8;
		let y = L.plotY + 4;
		for (const it of items) {
			root.appendChild(svg('rect', { x, y, width: 10, height: 10, fill: it.color }));
			root.appendChild(text(x + 14, y + 9, truncate(it.name, 14)));
			y += 16;
			if (y > L.plotY + L.plotH) break;
		}
	} else if (pos === 'b') {
		let x = L.plotX;
		const y = height - L.mb + 24;
		for (const it of items) {
			root.appendChild(svg('rect', { x, y: y - 9, width: 10, height: 10, fill: it.color }));
			const name = truncate(it.name, 16);
			root.appendChild(text(x + 14, y, name));
			x += 14 + name.length * 6 + 12;
			if (x > width - 8) break;
		}
	}
}

function drawTitle(root: SVGElement, model: ChartModel, width: number): void {
	if (!model.title) return;
	root.appendChild(text(width / 2, 20, model.title, { 'font-size': 14, 'text-anchor': 'middle', fill: '#222' }));
}

// Draw x/y axis lines, y-tick labels, and x-category labels (centered under
// each slot). Y ticks come from niceTicks over `domain`.
function drawAxes(
	root: SVGElement,
	L: Layout,
	domain: { min: number; max: number },
	categories: string[],
	percent: boolean,
): void {
	// Axis lines.
	root.appendChild(svg('line', { x1: L.plotX, y1: L.plotY, x2: L.plotX, y2: L.plotY + L.plotH, stroke: '#888' }));
	root.appendChild(svg('line', { x1: L.plotX, y1: L.plotY + L.plotH, x2: L.plotX + L.plotW, y2: L.plotY + L.plotH, stroke: '#888' }));

	// Y ticks.
	const ticks = niceTicks(domain.min, domain.max, 5);
	const yMin = ticks[0];
	const yMax = ticks[ticks.length - 1];
	const ySpan = yMax - yMin || 1;
	for (const t of ticks) {
		const y = L.plotY + L.plotH - ((t - yMin) / ySpan) * L.plotH;
		root.appendChild(svg('line', { x1: L.plotX - 3, y1: y, x2: L.plotX, y2: y, stroke: '#888' }));
		const label = percent ? `${Math.round(t)}%` : formatNum(t);
		root.appendChild(text(L.plotX - 6, y + 3, label, { 'text-anchor': 'end' }));
	}

	// X category ticks.
	const n = categories.length;
	if (n > 0) {
		const slot = L.plotW / n;
		const maxChars = Math.max(1, Math.floor((slot * 0.9) / 6)); // rough 6px/char
		for (let i = 0; i < n; i++) {
			const cx = L.plotX + slot * (i + 0.5);
			root.appendChild(svg('line', { x1: cx, y1: L.plotY + L.plotH, x2: cx, y2: L.plotY + L.plotH + 3, stroke: '#888' }));
			root.appendChild(text(cx, L.plotY + L.plotH + 14, truncate(categories[i], maxChars), { 'text-anchor': 'middle' }));
		}
	}
}

function formatNum(v: number): string {
	if (!Number.isFinite(v)) return '';
	if (Math.abs(v) >= 1000) return v.toLocaleString();
	if (Number.isInteger(v)) return String(v);
	return v.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
}

function drawColumn(root: SVGElement, model: ChartModel, L: Layout): void {
	const n = model.categories.length;
	if (n === 0) return;
	const percent = model.grouping === 'percentStacked';
	const domain = valueDomain(model);
	const ticks = niceTicks(domain.min, domain.max, 5);
	const yMin = ticks[0], yMax = ticks[ticks.length - 1];
	const ySpan = yMax - yMin || 1;
	const toY = (v: number) => L.plotY + L.plotH - ((v - yMin) / ySpan) * L.plotH;
	const slot = L.plotW / n;

	drawAxes(root, L, { min: yMin, max: yMax }, model.categories, percent);

	// Zero line when domain crosses zero — useful for negative stacks.
	if (yMin < 0 && yMax > 0) {
		const zy = toY(0);
		root.appendChild(svg('line', { x1: L.plotX, y1: zy, x2: L.plotX + L.plotW, y2: zy, stroke: '#888' }));
	}

	if (model.grouping === 'clustered' || model.grouping === 'standard') {
		const sn = Math.max(1, model.series.length);
		const groupW = slot * 0.8;
		const barW = groupW / sn;
		const groupX0 = (i: number) => L.plotX + slot * i + (slot - groupW) / 2;
		for (let si = 0; si < model.series.length; si++) {
			const s = model.series[si];
			const color = seriesColor(s, si);
			for (let i = 0; i < n; i++) {
				const v = s.values[i] ?? 0;
				const x = groupX0(i) + si * barW;
				const y0 = toY(0);
				const y1 = toY(v);
				const y = Math.min(y0, y1);
				const h = Math.abs(y1 - y0);
				root.appendChild(svg('rect', { x, y, width: Math.max(0.5, barW - 1), height: Math.max(0, h), fill: color }));
			}
		}
	} else {
		// stacked / percentStacked: single bar per slot, segments stacked.
		const barW = slot * 0.7;
		for (let i = 0; i < n; i++) {
			let total = 0;
			if (percent) {
				for (const s of model.series) total += Math.max(0, s.values[i] ?? 0);
				if (total === 0) total = 1;
			}
			let posAcc = 0, negAcc = 0;
			for (let si = 0; si < model.series.length; si++) {
				const s = model.series[si];
				let v = s.values[i] ?? 0;
				if (percent) v = (Math.max(0, v) / total) * 100;
				const color = seriesColor(s, si);
				const x = L.plotX + slot * i + (slot - barW) / 2;
				if (v >= 0) {
					const yTop = toY(posAcc + v);
					const yBot = toY(posAcc);
					root.appendChild(svg('rect', { x, y: yTop, width: barW, height: Math.max(0, yBot - yTop), fill: color }));
					posAcc += v;
				} else {
					const yTop = toY(negAcc);
					const yBot = toY(negAcc + v);
					root.appendChild(svg('rect', { x, y: yTop, width: barW, height: Math.max(0, yBot - yTop), fill: color }));
					negAcc += v;
				}
			}
		}
	}
}

function drawBar(root: SVGElement, model: ChartModel, L: Layout): void {
	// Horizontal bars: categories down the y-axis, values on x-axis.
	const n = model.categories.length;
	if (n === 0) return;
	const percent = model.grouping === 'percentStacked';
	const domain = valueDomain(model);
	const ticks = niceTicks(domain.min, domain.max, 5);
	const xMin = ticks[0], xMax = ticks[ticks.length - 1];
	const xSpan = xMax - xMin || 1;
	const toX = (v: number) => L.plotX + ((v - xMin) / xSpan) * L.plotW;
	const slot = L.plotH / n;

	// Axis lines.
	root.appendChild(svg('line', { x1: L.plotX, y1: L.plotY, x2: L.plotX, y2: L.plotY + L.plotH, stroke: '#888' }));
	root.appendChild(svg('line', { x1: L.plotX, y1: L.plotY + L.plotH, x2: L.plotX + L.plotW, y2: L.plotY + L.plotH, stroke: '#888' }));
	// X ticks (values).
	for (const t of ticks) {
		const x = toX(t);
		root.appendChild(svg('line', { x1: x, y1: L.plotY + L.plotH, x2: x, y2: L.plotY + L.plotH + 3, stroke: '#888' }));
		const label = percent ? `${Math.round(t)}%` : formatNum(t);
		root.appendChild(text(x, L.plotY + L.plotH + 14, label, { 'text-anchor': 'middle' }));
	}
	// Y ticks (categories).
	const maxChars = Math.max(1, Math.floor((L.ml - 8) / 6));
	for (let i = 0; i < n; i++) {
		const cy = L.plotY + slot * (i + 0.5);
		root.appendChild(svg('line', { x1: L.plotX - 3, y1: cy, x2: L.plotX, y2: cy, stroke: '#888' }));
		root.appendChild(text(L.plotX - 6, cy + 3, truncate(model.categories[i], maxChars), { 'text-anchor': 'end' }));
	}

	if (xMin < 0 && xMax > 0) {
		const zx = toX(0);
		root.appendChild(svg('line', { x1: zx, y1: L.plotY, x2: zx, y2: L.plotY + L.plotH, stroke: '#888' }));
	}

	if (model.grouping === 'clustered' || model.grouping === 'standard') {
		const sn = Math.max(1, model.series.length);
		const groupH = slot * 0.8;
		const barH = groupH / sn;
		for (let si = 0; si < model.series.length; si++) {
			const s = model.series[si];
			const color = seriesColor(s, si);
			for (let i = 0; i < n; i++) {
				const v = s.values[i] ?? 0;
				const y = L.plotY + slot * i + (slot - groupH) / 2 + si * barH;
				const x0 = toX(0);
				const x1 = toX(v);
				const x = Math.min(x0, x1);
				const w = Math.abs(x1 - x0);
				root.appendChild(svg('rect', { x, y, width: Math.max(0, w), height: Math.max(0.5, barH - 1), fill: color }));
			}
		}
	} else {
		const barH = slot * 0.7;
		for (let i = 0; i < n; i++) {
			let total = 0;
			if (percent) {
				for (const s of model.series) total += Math.max(0, s.values[i] ?? 0);
				if (total === 0) total = 1;
			}
			let posAcc = 0, negAcc = 0;
			for (let si = 0; si < model.series.length; si++) {
				const s = model.series[si];
				let v = s.values[i] ?? 0;
				if (percent) v = (Math.max(0, v) / total) * 100;
				const color = seriesColor(s, si);
				const y = L.plotY + slot * i + (slot - barH) / 2;
				if (v >= 0) {
					const x0 = toX(posAcc);
					const x1 = toX(posAcc + v);
					root.appendChild(svg('rect', { x: x0, y, width: Math.max(0, x1 - x0), height: barH, fill: color }));
					posAcc += v;
				} else {
					const x0 = toX(negAcc + v);
					const x1 = toX(negAcc);
					root.appendChild(svg('rect', { x: x0, y, width: Math.max(0, x1 - x0), height: barH, fill: color }));
					negAcc += v;
				}
			}
		}
	}
}

function drawLine(root: SVGElement, model: ChartModel, L: Layout): void {
	const n = model.categories.length;
	if (n === 0) return;
	const domain = valueDomain(model);
	const ticks = niceTicks(domain.min, domain.max, 5);
	const yMin = ticks[0], yMax = ticks[ticks.length - 1];
	const ySpan = yMax - yMin || 1;
	const toY = (v: number) => L.plotY + L.plotH - ((v - yMin) / ySpan) * L.plotH;
	const slot = L.plotW / n;

	drawAxes(root, L, { min: yMin, max: yMax }, model.categories, false);

	for (let si = 0; si < model.series.length; si++) {
		const s = model.series[si];
		const color = seriesColor(s, si);
		const pts: string[] = [];
		for (let i = 0; i < n; i++) {
			const cx = L.plotX + slot * (i + 0.5);
			const cy = toY(s.values[i] ?? 0);
			pts.push(`${cx},${cy}`);
		}
		root.appendChild(svg('polyline', {
			points: pts.join(' '),
			fill: 'none',
			stroke: color,
			'stroke-width': 2,
		}));
	}
}

// Shared arc path builder for pie/doughnut slices. For pie, innerR = 0 gives
// M cx cy L x1 y1 A r r 0 largeArc 1 x2 y2 Z. For doughnut we draw the outer
// arc and then the inner arc back the other way, closing with Z.
function arcPath(cx: number, cy: number, r: number, innerR: number, a0: number, a1: number): string {
	const large = (a1 - a0) > Math.PI ? 1 : 0;
	const x1 = cx + r * Math.cos(a0), y1 = cy + r * Math.sin(a0);
	const x2 = cx + r * Math.cos(a1), y2 = cy + r * Math.sin(a1);
	if (innerR <= 0) {
		return `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`;
	}
	const xi1 = cx + innerR * Math.cos(a1), yi1 = cy + innerR * Math.sin(a1);
	const xi2 = cx + innerR * Math.cos(a0), yi2 = cy + innerR * Math.sin(a0);
	return `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${xi1} ${yi1} A ${innerR} ${innerR} 0 ${large} 0 ${xi2} ${yi2} Z`;
}

function drawPieOrDoughnut(root: SVGElement, model: ChartModel, L: Layout, doughnut: boolean): void {
	// Pie shows series[0].values as slices keyed by category.
	const vals = model.series[0]?.values ?? [];
	const total = vals.reduce((a, b) => a + Math.max(0, b), 0);
	if (total <= 0) return;
	const cx = L.plotX + L.plotW / 2;
	const cy = L.plotY + L.plotH / 2;
	const r = Math.min(L.plotW, L.plotH) / 2 - 4;
	const innerR = doughnut ? r * 0.5 : 0;
	let a = -Math.PI / 2;
	for (let i = 0; i < vals.length; i++) {
		const v = Math.max(0, vals[i]);
		if (v === 0) continue;
		const aNext = a + (v / total) * Math.PI * 2;
		const color = PALETTE[i % PALETTE.length];
		const d = arcPath(cx, cy, r, innerR, a, aNext);
		root.appendChild(svg('path', { d, fill: color, stroke: '#fff', 'stroke-width': 1 }));
		a = aNext;
	}
}

// For pie / doughnut we want the legend to enumerate categories, not the
// single-series name — synthesise a category-flavored model for drawLegend.
function pieLegendModel(model: ChartModel): ChartModel {
	const vals = model.series[0]?.values ?? [];
	const series: ChartSeries[] = vals.map((_, i) => ({
		name: model.categories[i] || `Slice ${i + 1}`,
		colorHex: PALETTE[i % PALETTE.length],
		values: [],
	}));
	return { ...model, series };
}

export function renderChart(model: ChartModel, widthPx: number, heightPx: number): SVGSVGElement {
	const width = Math.max(60, Math.round(widthPx));
	const height = Math.max(60, Math.round(heightPx));
	const root = svg('svg', {
		xmlns: SVG_NS,
		width,
		height,
		viewBox: `0 0 ${width} ${height}`,
	}) as SVGSVGElement;

	const L = layout(model, width, height);
	drawTitle(root, model, width);

	switch (model.kind) {
		case 'column':
			drawColumn(root, model, L);
			drawLegend(root, model, L, width, height);
			break;
		case 'bar':
			drawBar(root, model, L);
			drawLegend(root, model, L, width, height);
			break;
		case 'line':
			drawLine(root, model, L);
			drawLegend(root, model, L, width, height);
			break;
		case 'pie':
			drawPieOrDoughnut(root, model, L, false);
			drawLegend(root, pieLegendModel(model), L, width, height);
			break;
		case 'doughnut':
			drawPieOrDoughnut(root, model, L, true);
			drawLegend(root, pieLegendModel(model), L, width, height);
			break;
	}

	return root;
}
