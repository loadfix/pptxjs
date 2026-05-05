// Loader for ppt/tableStyles.xml.
//
// The part contains a <a:tblStyleLst> with <a:tblStyle styleId="{UUID}">
// entries. Each style declares a set of "band" definitions (wholeTbl,
// firstRow, lastRow, band1H/band2H for row banding, band1V/band2V for
// column banding, firstCol, lastCol, nwCell, neCell, swCell, seCell).
// Each band has an <a:tcStyle> child carrying per-side borders
// (<a:tcBdr>) and a fill, plus an optional <a:tcTxStyle> for text colour.
//
// We expose the parsed styles as a Map<styleId, TableStyle>. If the
// referenced tableStyles part is absent, or a slide's styleId has no
// matching entry, callers should fall back to defaults (no whole-table
// chrome, render per-cell XML only).

import { A_NS } from './namespaces';
import { firstChildNS, childrenNS } from './xml-utils';
import { OpenXmlPackage, resolveRelTarget } from './open-xml-package';
import { LineStyle, Fill, SolidFill, parseLine, parseFillElement } from './fill';
import { ThemeColors, ClrMap, emptyThemeColors, emptyClrMap, parseTheme, parseClrMap, resolveColorElement } from './theme';

const TABLE_STYLES_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles";
const THEME_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";

// Per-side borders (left/right/top/bottom) plus the two diagonals
// (tl-br and bl-tr). All fields are optional — inheritance across bands
// merges by field.
export interface TableBorders {
	l: LineStyle | null;
	r: LineStyle | null;
	t: LineStyle | null;
	b: LineStyle | null;
	insideH: LineStyle | null;
	insideV: LineStyle | null;
	tlbr: LineStyle | null;
	blTr: LineStyle | null;
}

export interface TableBandStyle {
	borders: TableBorders;
	fill: Fill | null;
	textColorHex: string | null;
}

// A TableStyle captures enough of the band definitions for the renderer
// to cascade: whole-table defaults first, then row-oriented overrides
// (band1H/band2H/firstRow/lastRow), column banding (band1V/band2V with
// firstCol/lastCol overrides), and finally the four corner cells
// (nwCell/neCell/swCell/seCell) which win when both surrounding row and
// column flags are set.
export interface TableStyle {
	styleId: string;
	wholeTbl: TableBandStyle;
	firstRow: TableBandStyle;
	lastRow: TableBandStyle;
	firstCol: TableBandStyle;
	lastCol: TableBandStyle;
	band1H: TableBandStyle;
	band2H: TableBandStyle;
	band1V: TableBandStyle;
	band2V: TableBandStyle;
	nwCell: TableBandStyle;
	neCell: TableBandStyle;
	swCell: TableBandStyle;
	seCell: TableBandStyle;
}

export function emptyTableBorders(): TableBorders {
	return { l: null, r: null, t: null, b: null, insideH: null, insideV: null, tlbr: null, blTr: null };
}

export function emptyBand(): TableBandStyle {
	return { borders: emptyTableBorders(), fill: null, textColorHex: null };
}

function emptyStyle(styleId: string): TableStyle {
	return {
		styleId,
		wholeTbl: emptyBand(),
		firstRow: emptyBand(),
		lastRow: emptyBand(),
		firstCol: emptyBand(),
		lastCol: emptyBand(),
		band1H: emptyBand(),
		band2H: emptyBand(),
		band1V: emptyBand(),
		band2V: emptyBand(),
		nwCell: emptyBand(),
		neCell: emptyBand(),
		swCell: emptyBand(),
		seCell: emptyBand(),
	};
}

// Locate the tableStyles part via presentation.xml relationships. Returns
// null when the part isn't referenced (e.g. the presentation never had a
// styled table).
async function findTableStylesPath(pkg: OpenXmlPackage): Promise<string | null> {
	const presPath = "ppt/presentation.xml";
	const rels = await pkg.loadRelationships(presPath);
	for (const r of rels.values()) {
		if (r.type === TABLE_STYLES_REL_TYPE) return resolveRelTarget(presPath, r.target);
	}
	return null;
}

// Resolve a theme/clrMap pair for colour resolution inside tableStyles.
// Table styles reference scheme colours (e.g. "accent1"), so we need the
// master's clrMap and the associated theme. We use the first master/theme
// found in the presentation — in practice presentations use a consistent
// colour scheme for their table styles, and per-master table styles are
// extremely rare.
async function findClrMapAndTheme(pkg: OpenXmlPackage): Promise<{ clrMap: ClrMap; theme: ThemeColors }> {
	const presPath = "ppt/presentation.xml";
	const rels = await pkg.loadRelationships(presPath);
	for (const r of rels.values()) {
		if (r.type !== MASTER_REL_TYPE) continue;
		const masterPath = resolveRelTarget(presPath, r.target);
		const masterDoc = await pkg.loadXml(masterPath);
		if (!masterDoc) continue;
		const clrMap = parseClrMap(masterDoc);
		const masterRels = await pkg.loadRelationships(masterPath);
		for (const mr of masterRels.values()) {
			if (mr.type !== THEME_REL_TYPE) continue;
			const themePath = resolveRelTarget(masterPath, mr.target);
			const themeDoc = await pkg.loadXml(themePath);
			if (!themeDoc) break;
			return { clrMap, theme: parseTheme(themeDoc) };
		}
		return { clrMap, theme: emptyThemeColors() };
	}
	return { clrMap: emptyClrMap(), theme: emptyThemeColors() };
}

function parseTcBdr(tcBdr: Element | null, clrMap: ClrMap, theme: ThemeColors): TableBorders {
	const out = emptyTableBorders();
	if (!tcBdr) return out;
	const side = (localName: string): LineStyle | null => {
		const holder = firstChildNS(tcBdr, A_NS.a, localName);
		if (!holder) return null;
		const ln = firstChildNS(holder, A_NS.a, "ln");
		return parseLine(ln, clrMap, theme);
	};
	out.l = side("left");
	out.r = side("right");
	out.t = side("top");
	out.b = side("bottom");
	out.insideH = side("insideH");
	out.insideV = side("insideV");
	out.tlbr = side("tl2br");
	out.blTr = side("tr2bl");
	return out;
}

// tcStyle children: <a:tcBdr>, fill (<a:fill><a:solidFill/.../></a:fill>
// wrapper or a direct fill element), and a <a:tcTxStyle> with a colour.
function parseTcStyle(tcStyle: Element | null, clrMap: ClrMap, theme: ThemeColors): TableBandStyle {
	const band = emptyBand();
	if (!tcStyle) return band;
	band.borders = parseTcBdr(firstChildNS(tcStyle, A_NS.a, "tcBdr"), clrMap, theme);

	// Fill: tableStyles uses an <a:fill> wrapper that contains one of the
	// fill kinds. Handle both the wrapped and unwrapped forms.
	const fillWrap = firstChildNS(tcStyle, A_NS.a, "fill");
	if (fillWrap) {
		for (const child of Array.from(fillWrap.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			band.fill = parseFillElement(child, clrMap, theme);
			if (band.fill) break;
		}
	} else {
		for (const child of Array.from(tcStyle.children)) {
			if (child.namespaceURI !== A_NS.a) continue;
			const ln = child.localName;
			if (ln === 'solidFill' || ln === 'noFill' || ln === 'gradFill' || ln === 'blipFill' || ln === 'pattFill') {
				band.fill = parseFillElement(child, clrMap, theme);
				break;
			}
		}
	}
	return band;
}

function parseTcTxStyle(tcTxStyle: Element | null, clrMap: ClrMap, theme: ThemeColors): string | null {
	if (!tcTxStyle) return null;
	// Any color child (srgbClr/schemeClr/...) resolves via the usual path.
	for (const c of Array.from(tcTxStyle.children)) {
		if (c.namespaceURI !== A_NS.a) continue;
		const ln = c.localName;
		if (
			ln === 'srgbClr' || ln === 'schemeClr' || ln === 'prstClr'
			|| ln === 'sysClr' || ln === 'scrgbClr' || ln === 'hslClr'
		) {
			const r = resolveColorElement(c, clrMap, theme);
			return r?.colorHex ?? null;
		}
	}
	return null;
}

function parseBand(bandEl: Element | null, clrMap: ClrMap, theme: ThemeColors): TableBandStyle {
	if (!bandEl) return emptyBand();
	const tcStyle = firstChildNS(bandEl, A_NS.a, "tcStyle");
	const band = parseTcStyle(tcStyle, clrMap, theme);
	const tcTxStyle = firstChildNS(bandEl, A_NS.a, "tcTxStyle");
	band.textColorHex = parseTcTxStyle(tcTxStyle, clrMap, theme);
	return band;
}

function parseTableStyle(el: Element, clrMap: ClrMap, theme: ThemeColors): TableStyle | null {
	const styleId = el.getAttribute("styleId");
	if (!styleId) return null;
	const style = emptyStyle(styleId);

	// The DrawingML spec uses <a:tblBg> + <a:wholeTbl>/<a:band1H>/... as
	// direct children of <a:tblStyle>. Walk children and route by localName.
	for (const c of childrenNS(el, A_NS.a, "wholeTbl")) style.wholeTbl = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "firstRow")) style.firstRow = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "lastRow")) style.lastRow = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "firstCol")) style.firstCol = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "lastCol")) style.lastCol = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "band1H")) style.band1H = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "band2H")) style.band2H = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "band1V")) style.band1V = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "band2V")) style.band2V = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "nwCell")) style.nwCell = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "neCell")) style.neCell = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "swCell")) style.swCell = parseBand(c, clrMap, theme);
	for (const c of childrenNS(el, A_NS.a, "seCell")) style.seCell = parseBand(c, clrMap, theme);

	return style;
}

export async function loadTableStyles(pkg: OpenXmlPackage): Promise<Map<string, TableStyle>> {
	const out = new Map<string, TableStyle>();
	const path = await findTableStylesPath(pkg);
	if (!path) return out;
	const doc = await pkg.loadXml(path);
	if (!doc) return out;

	const { clrMap, theme } = await findClrMapAndTheme(pkg);

	const root = doc.documentElement;
	// Root is <a:tblStyleLst>; children are <a:tblStyle>.
	for (const c of childrenNS(root, A_NS.a, "tblStyle")) {
		const style = parseTableStyle(c, clrMap, theme);
		if (style) out.set(style.styleId, style);
	}
	return out;
}

// Re-export the SolidFill type so downstream consumers (renderer) can
// type-guard without importing fill.ts directly. Kept minimal on purpose.
export type { SolidFill };
