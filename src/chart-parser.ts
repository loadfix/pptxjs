// Parse a ppt/charts/chartN.xml Document into a minimal ChartModel suitable
// for the inline SVG renderer. The parser handles the five chart families
// we render natively (bar/column/line/pie/doughnut); anything else returns
// null and leaves the "[Chart]" placeholder in place.
//
// The reader stays deliberately tolerant: missing attributes fall back to
// OOXML defaults, theme-color fills (schemeClr) degrade to null so the
// renderer can paint from its own palette, and value arrays are padded
// with zeros at missing indices so index-sparse caches render cleanly.

const C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

export type ChartKind = 'bar' | 'column' | 'line' | 'pie' | 'doughnut';
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';
export type ChartLegendPos = 'r' | 'b' | 't' | 'l' | 'tr' | null;

export interface ChartSeries {
	name: string;
	colorHex: string | null;
	values: number[];
}

export interface ChartModel {
	kind: ChartKind;
	grouping: ChartGrouping;
	series: ChartSeries[];
	categories: string[];
	title: string | null;
	legend: ChartLegendPos;
}

function firstC(parent: Element | null, local: string): Element | null {
	if (!parent) return null;
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === C_NS && c.localName === local) return c;
	}
	return null;
}

function allC(parent: Element | null, local: string): Element[] {
	if (!parent) return [];
	const out: Element[] = [];
	for (const c of Array.from(parent.children)) {
		if (c.namespaceURI === C_NS && c.localName === local) out.push(c);
	}
	return out;
}

function attrVal(el: Element | null): string | null {
	return el?.getAttribute('val') ?? null;
}

// Extract the series-ordered string cache from a <c:strRef> or <c:numRef>.
// Returns an index-sparse map from point index to value, along with the
// declared ptCount (when present) so callers know how many slots to fill.
function readCache(ref: Element | null, isNum: boolean): { count: number; data: Map<number, string> } {
	const data = new Map<number, string>();
	if (!ref) return { count: 0, data };
	const cache = firstC(ref, isNum ? 'numCache' : 'strCache');
	if (!cache) return { count: 0, data };
	const countAttr = attrVal(firstC(cache, 'ptCount'));
	const count = countAttr != null ? Number(countAttr) : 0;
	for (const pt of allC(cache, 'pt')) {
		const idx = Number(pt.getAttribute('idx') ?? '0');
		const v = firstC(pt, 'v');
		if (v && v.textContent != null) data.set(idx, v.textContent);
	}
	return { count, data };
}

function readSeriesName(ser: Element): string {
	const tx = firstC(ser, 'tx');
	if (!tx) return '';
	const strRef = firstC(tx, 'strRef');
	if (strRef) {
		const { data } = readCache(strRef, false);
		// strRef with a single cached point: take the first entry (usually idx 0).
		for (const v of data.values()) return v;
		return '';
	}
	// Fallback: <c:tx><c:v>text</c:v>.
	const v = firstC(tx, 'v');
	if (v && v.textContent != null) return v.textContent;
	// Or plain strings on tx directly (rare).
	return tx.textContent?.trim() ?? '';
}

function readSeriesColor(ser: Element): string | null {
	const spPr = firstC(ser, 'spPr');
	if (!spPr) return null;
	// <a:solidFill><a:srgbClr val="RRGGBB"/></a:solidFill>
	for (const c of Array.from(spPr.children)) {
		if (c.namespaceURI === A_NS && c.localName === 'solidFill') {
			for (const inner of Array.from(c.children)) {
				if (inner.namespaceURI === A_NS && inner.localName === 'srgbClr') {
					const v = inner.getAttribute('val');
					return v ? `#${v}` : null;
				}
			}
			// schemeClr or other: treat as theme-driven, let renderer pick.
			return null;
		}
	}
	return null;
}

function readSeriesValues(ser: Element): number[] {
	const val = firstC(ser, 'val');
	if (!val) return [];
	const numRef = firstC(val, 'numRef');
	const { count, data } = readCache(numRef, true);
	const size = Math.max(count, ...Array.from(data.keys(), (k) => k + 1), 0);
	const out: number[] = new Array(size).fill(0);
	for (const [idx, v] of data) {
		const n = Number(v);
		out[idx] = Number.isFinite(n) ? n : 0;
	}
	return out;
}

function readCategories(ser: Element): string[] {
	const cat = firstC(ser, 'cat');
	if (!cat) return [];
	const strRef = firstC(cat, 'strRef');
	if (strRef) {
		const { count, data } = readCache(strRef, false);
		const size = Math.max(count, ...Array.from(data.keys(), (k) => k + 1), 0);
		const out: string[] = new Array(size).fill('');
		for (const [idx, v] of data) out[idx] = v;
		return out;
	}
	const numRef = firstC(cat, 'numRef');
	if (numRef) {
		const { count, data } = readCache(numRef, true);
		const size = Math.max(count, ...Array.from(data.keys(), (k) => k + 1), 0);
		const out: string[] = new Array(size).fill('');
		for (const [idx, v] of data) out[idx] = v;
		return out;
	}
	return [];
}

function readTitle(chart: Element | null): string | null {
	const title = firstC(chart, 'title');
	if (!title) return null;
	const tx = firstC(title, 'tx');
	if (!tx) return null;
	const rich = firstC(tx, 'rich');
	if (!rich) return null;
	// Concatenate all <a:t> descendants in document order.
	const texts = rich.getElementsByTagNameNS(A_NS, 't');
	let s = '';
	for (let i = 0; i < texts.length; i++) s += texts[i].textContent ?? '';
	return s.length > 0 ? s : null;
}

function readLegend(chart: Element | null): ChartLegendPos {
	const legend = firstC(chart, 'legend');
	if (!legend) return null;
	const pos = attrVal(firstC(legend, 'legendPos'));
	if (pos === 'r' || pos === 'b' || pos === 't' || pos === 'l' || pos === 'tr') return pos;
	return 'r'; // ECMA default when legend present
}

function groupingFrom(plot: Element): ChartGrouping {
	const v = attrVal(firstC(plot, 'grouping'));
	if (v === 'clustered' || v === 'stacked' || v === 'percentStacked') return v;
	return 'standard';
}

export function parseChart(doc: Document): ChartModel | null {
	const root = doc.documentElement;
	if (!root || root.namespaceURI !== C_NS || root.localName !== 'chartSpace') return null;
	const chart = firstC(root, 'chart');
	if (!chart) return null;
	const plotArea = firstC(chart, 'plotArea');
	if (!plotArea) return null;

	// Find first supported plot-area chart element.
	let plot: Element | null = null;
	let kind: ChartKind | null = null;
	for (const c of Array.from(plotArea.children)) {
		if (c.namespaceURI !== C_NS) continue;
		const ln = c.localName;
		if (ln === 'barChart') {
			const dir = attrVal(firstC(c, 'barDir'));
			kind = dir === 'bar' ? 'bar' : 'column';
			plot = c;
			break;
		}
		if (ln === 'lineChart') { kind = 'line'; plot = c; break; }
		if (ln === 'pieChart') { kind = 'pie'; plot = c; break; }
		if (ln === 'doughnutChart') { kind = 'doughnut'; plot = c; break; }
	}
	if (!plot || !kind) return null;

	const grouping: ChartGrouping =
		kind === 'pie' || kind === 'doughnut' ? 'standard' : groupingFrom(plot);

	let serEls = allC(plot, 'ser');
	if (kind === 'pie' && serEls.length > 1) serEls = serEls.slice(0, 1);

	const series: ChartSeries[] = serEls.map((s) => ({
		name: readSeriesName(s),
		colorHex: readSeriesColor(s),
		values: readSeriesValues(s),
	}));

	// Pad all series to the same length so the renderer can index uniformly.
	const maxLen = series.reduce((m, s) => Math.max(m, s.values.length), 0);
	for (const s of series) {
		while (s.values.length < maxLen) s.values.push(0);
	}

	// Categories shared across series — pull from the first series that has them.
	let categories: string[] = [];
	for (const s of serEls) {
		const cats = readCategories(s);
		if (cats.length > 0) { categories = cats; break; }
	}
	// If categories came up short, pad with empty strings to maxLen.
	while (categories.length < maxLen) categories.push('');

	return {
		kind,
		grouping,
		series,
		categories,
		title: readTitle(chart),
		legend: readLegend(chart),
	};
}
