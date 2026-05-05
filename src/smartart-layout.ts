// Minimal SmartArt layout engine fallback.
//
// A full OOXML SmartArt layout engine (ECMA-376 Part 1 §21.4) handles ~20
// categories of algorithms with many parameters each and is a multi-month
// project. For pptxjs we only cover two of the most common templates and
// use them when the pre-rendered `diagrams/drawingN.xml` cache is absent:
//
//   - Linear list (horizontal / vertical): hList1, vList1, and a handful
//     of closely-related variants. Each top-level node becomes a flat box
//     in a row / column.
//   - Simple tree / hierarchy: hierRoot1 and close relatives. Each level
//     of the hierarchy is a horizontal row of boxes; parents center over
//     their children and a stroke connects them.
//
// Anything else returns `{ algorithm: 'unknown' }` and the caller keeps
// the existing `[SmartArt]` placeholder — trying to guess a layout for
// cycles, matrices, radial templates, etc. produces worse output than
// the placeholder.
//
// Positions are expressed in EMUs (the unit used everywhere else in
// pptxjs) so the synthesised shapes slot straight into the slide's shape
// list. Colors are pulled from the theme's accent1 where possible so the
// fallback at least looks related to the deck; no per-node styling is
// honoured.

import { OpenXmlPackage } from './open-xml-package';
import { firstChildNS, childrenNS } from './xml-utils';
import { A_NS } from './namespaces';
import type { ShapeLike } from './presentation-parser';
import type { ThemeColors } from './theme';

// DrawingML diagram namespace — the one used in `diagrams/data1.xml` and
// `diagrams/layout1.xml`. Not the `dsp` namespace used by the rendered
// drawing cache.
const DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

// Parsed data-model node. `id` is the `<dgm:pt modelId="…">` attribute;
// `text` is the concatenated plain text of all runs in `<dgm:t>` (preserving
// paragraph boundaries with a single `\n`). `children` are in the order
// they appear in the connection list (sorted by `destOrd`).
export interface DgmNode {
	id: string;
	text: string;
	children: DgmNode[];
}

export type DgmAlgorithm = 'hList' | 'vList' | 'hier' | 'unknown';

interface RawPt {
	id: string;
	type: string; // node / asst / doc / pres / parTrans / sibTrans
	text: string;
}

interface RawCxn {
	srcId: string;
	destId: string;
	srcOrd: number;
	destOrd: number;
	type: string; // parOf / presOf / presParOf / unknownRelationship
}

// Parse `diagrams/dataN.xml` into a single root node. The OOXML data model
// is a list of `<dgm:pt>` elements plus a list of `<dgm:cxn>` edges. We
// use `type="parOf"` edges as the parent-child relationship and pick the
// `type="doc"` point as the synthetic root (falling back to whichever
// point has no parent).
//
// Non-data points (`type="pres"`, `parTrans`, `sibTrans`) are ignored —
// these carry presentation-layer connector state we don't reproduce.
export async function parseDgmData(
	pkg: OpenXmlPackage,
	dataPath: string,
): Promise<{ root: DgmNode | null }> {
	const doc = await pkg.loadXml(dataPath);
	if (!doc) return { root: null };
	const root = doc.documentElement;
	if (!root || root.namespaceURI !== DGM_NS || root.localName !== "dataModel") {
		return { root: null };
	}

	const ptLst = firstChildNS(root, DGM_NS, "ptLst");
	if (!ptLst) return { root: null };

	const rawPts: RawPt[] = [];
	for (const ptEl of childrenNS(ptLst, DGM_NS, "pt")) {
		const id = ptEl.getAttribute("modelId") ?? "";
		if (!id) continue;
		const type = ptEl.getAttribute("type") ?? "node";
		const tEl = firstChildNS(ptEl, DGM_NS, "t");
		const text = tEl ? extractTextBody(tEl) : "";
		rawPts.push({ id, type, text });
	}

	const cxnLst = firstChildNS(root, DGM_NS, "cxnLst");
	const rawCxns: RawCxn[] = [];
	if (cxnLst) {
		for (const cxn of childrenNS(cxnLst, DGM_NS, "cxn")) {
			const srcId = cxn.getAttribute("srcId") ?? "";
			const destId = cxn.getAttribute("destId") ?? "";
			if (!srcId || !destId) continue;
			rawCxns.push({
				srcId,
				destId,
				srcOrd: Number(cxn.getAttribute("srcOrd")) || 0,
				destOrd: Number(cxn.getAttribute("destOrd")) || 0,
				type: cxn.getAttribute("type") ?? "parOf",
			});
		}
	}

	// Keep only data-bearing points. `pres`, `parTrans`, `sibTrans` are
	// layout scaffolding; `doc` is the synthetic document root.
	const dataPts = rawPts.filter((p) => p.type === "node" || p.type === "asst" || p.type === "doc");
	const ptMap = new Map<string, DgmNode>();
	for (const p of dataPts) {
		ptMap.set(p.id, { id: p.id, text: p.text, children: [] });
	}

	// Build parent-child edges from parOf connections only.
	const childParent = new Map<string, string>();
	// Preserve author order via destOrd (ECMA-376 says the destOrd of a
	// parOf edge is the child's position under the parent).
	type Edge = { srcId: string; destId: string; destOrd: number };
	const edges: Edge[] = [];
	for (const c of rawCxns) {
		if (c.type !== "parOf") continue;
		if (!ptMap.has(c.srcId) || !ptMap.has(c.destId)) continue;
		edges.push({ srcId: c.srcId, destId: c.destId, destOrd: c.destOrd });
		childParent.set(c.destId, c.srcId);
	}
	edges.sort((a, b) => a.destOrd - b.destOrd);
	for (const e of edges) {
		const parent = ptMap.get(e.srcId);
		const child = ptMap.get(e.destId);
		if (parent && child) parent.children.push(child);
	}

	// Root: prefer the `type="doc"` point; else whatever point has no
	// incoming parOf edge.
	const docPt = dataPts.find((p) => p.type === "doc");
	let rootNode: DgmNode | null = docPt ? ptMap.get(docPt.id) ?? null : null;
	if (!rootNode) {
		const orphan = dataPts.find((p) => !childParent.has(p.id));
		rootNode = orphan ? ptMap.get(orphan.id) ?? null : null;
	}
	return { root: rootNode };
}

// Pull the text content of a `<dgm:t>` element. The element follows the
// full DrawingML `a:txBody` schema; we ignore paragraph and run styling
// and just concatenate `<a:t>` contents with newlines between paragraphs.
function extractTextBody(tEl: Element): string {
	const paras: string[] = [];
	for (const p of childrenNS(tEl, A_NS.a, "p")) {
		const parts: string[] = [];
		for (const r of Array.from(p.children)) {
			if (r.namespaceURI !== A_NS.a) continue;
			if (r.localName === "r") {
				const tE = firstChildNS(r, A_NS.a, "t");
				if (tE) parts.push(tE.textContent ?? "");
			} else if (r.localName === "br") {
				parts.push("\n");
			}
		}
		paras.push(parts.join(""));
	}
	return paras.join("\n").trim();
}

// Inspect `<dgm:layoutDef uniqueId="…">` and classify the template.
// PowerPoint's built-in SmartArt templates all have uniqueIds of the
// form `urn:microsoft.com/office/officeart/2005/8/layout/<name>`. We
// pattern-match on `<name>` rather than the full URI so close variants
// (`hList2`, `vList3`) get the same treatment as the canonical template.
export async function parseDgmLayout(
	pkg: OpenXmlPackage,
	layoutPath: string,
): Promise<{ algorithm: DgmAlgorithm }> {
	const doc = await pkg.loadXml(layoutPath);
	if (!doc) return { algorithm: 'unknown' };
	const root = doc.documentElement;
	if (!root || root.namespaceURI !== DGM_NS || root.localName !== "layoutDef") {
		return { algorithm: 'unknown' };
	}
	const uniqueId = root.getAttribute("uniqueId") ?? "";
	// Extract the trailing name component, e.g.
	// "urn:microsoft.com/office/officeart/2005/8/layout/hList1" → "hList1".
	const slash = uniqueId.lastIndexOf("/");
	const name = slash >= 0 ? uniqueId.slice(slash + 1) : uniqueId;

	// Horizontal lists: hList1, hList2, ... plus horizontal-bullet variants.
	if (/^hList\d*$/.test(name) || name === "hChevron1" || name === "hChevron2") {
		return { algorithm: 'hList' };
	}
	// Vertical lists and close variants.
	if (/^vList\d*$/.test(name) || name === "lstBullet1" || name === "vertList1") {
		return { algorithm: 'vList' };
	}
	// Hierarchy family — hierRoot1, hierChild1, orgChart1, ... We only
	// render a plain top-down tree, which is a reasonable approximation
	// of all of these.
	if (/^hier/.test(name) || /^orgChart/.test(name)) {
		return { algorithm: 'hier' };
	}
	return { algorithm: 'unknown' };
}

// Synthesise shape boxes for a recognised layout. Returned shapes are
// expressed in the frame's local coordinate space (origin at 0,0). The
// caller is expected to offset by the frame's x/y like the drawing-cache
// expansion path does. `accentHex` is the theme accent color used for
// the fill; connector lines in the tree layout use a darker grey.
export function renderMinimalSmartArt(
	root: DgmNode | null,
	algorithm: DgmAlgorithm,
	frameCx: number,
	frameCy: number,
	theme: ThemeColors,
): ShapeLike[] {
	if (!root) return [];
	const accent = theme.scheme.accent1 ?? "4472C4";
	const strokeHex = "595959";

	if (algorithm === 'hList' || algorithm === 'vList') {
		// Top-level children of the (doc) root are the list items. If the
		// document root has no children (single-point diagram), treat the
		// root itself as the sole item so the user sees at least one box.
		const items = root.children.length > 0 ? root.children : [root];
		return layoutLinearList(items, algorithm, frameCx, frameCy, accent);
	}
	if (algorithm === 'hier') {
		// Hierarchy: the visible tree starts at the first child of the doc
		// root. A single-child doc (one top-level boss) becomes the real
		// root. A doc with multiple children (rare in PowerPoint's hier
		// templates but allowed by the schema) is rendered as a forest —
		// each child laid out as its own tree and concatenated.
		if (!root.text && root.children.length === 1) {
			return layoutHierarchy(root.children[0], frameCx, frameCy, accent, strokeHex);
		}
		if (!root.text && root.children.length > 1) {
			const out: ShapeLike[] = [];
			const per = Math.floor(frameCx / root.children.length);
			for (let i = 0; i < root.children.length; i++) {
				const sub = layoutHierarchy(root.children[i], per, frameCy, accent, strokeHex);
				for (const sh of sub) sh.x += i * per;
				out.push(...sub);
			}
			return out;
		}
		return layoutHierarchy(root, frameCx, frameCy, accent, strokeHex);
	}
	return [];
}

function layoutLinearList(
	items: DgmNode[],
	algorithm: 'hList' | 'vList',
	frameCx: number,
	frameCy: number,
	accentHex: string,
): ShapeLike[] {
	if (items.length === 0) return [];
	const n = items.length;
	// 5% gap between boxes on the laid-out axis, padding on the inside.
	const gapFrac = 0.05;
	const padFrac = 0.02;
	const out: ShapeLike[] = [];
	if (algorithm === 'hList') {
		const pad = Math.round(frameCx * padFrac);
		const avail = frameCx - 2 * pad;
		const totalGap = Math.round(avail * gapFrac * (n - 1));
		const boxW = Math.max(1, Math.floor((avail - totalGap) / n));
		const boxH = Math.max(1, frameCy - 2 * Math.round(frameCy * padFrac));
		const y = Math.round((frameCy - boxH) / 2);
		for (let i = 0; i < n; i++) {
			const x = pad + i * (boxW + Math.round(avail * gapFrac));
			out.push(makeBox(x, y, boxW, boxH, items[i].text, accentHex));
		}
	} else {
		const pad = Math.round(frameCy * padFrac);
		const avail = frameCy - 2 * pad;
		const totalGap = Math.round(avail * gapFrac * (n - 1));
		const boxH = Math.max(1, Math.floor((avail - totalGap) / n));
		const boxW = Math.max(1, frameCx - 2 * Math.round(frameCx * padFrac));
		const x = Math.round((frameCx - boxW) / 2);
		for (let i = 0; i < n; i++) {
			const y = pad + i * (boxH + Math.round(avail * gapFrac));
			out.push(makeBox(x, y, boxW, boxH, items[i].text, accentHex));
		}
	}
	return out;
}

// Simple layered tree layout:
//   1. Count leaves per subtree to budget horizontal slots.
//   2. Place leaves evenly across the frame, then set each non-leaf to
//      the horizontal centre of its children.
//   3. Vertical levels are evenly distributed.
//   4. Emit a thin connector line from each parent's bottom to each
//      child's top.
//
// Depth is measured from the tree root. Empty trees and trees with only
// a root node are handled.
function layoutHierarchy(
	root: DgmNode,
	frameCx: number,
	frameCy: number,
	accentHex: string,
	strokeHex: string,
): ShapeLike[] {
	// Flatten into depth-indexed lists.
	interface Placed {
		node: DgmNode;
		depth: number;
		parent: Placed | null;
		leafIdxStart: number;
		leafCount: number;
		cx: number;
		cy: number;
	}
	const nodes: Placed[] = [];
	const byDepth: Placed[][] = [];
	let leafCounter = 0;
	function visit(n: DgmNode, depth: number, parent: Placed | null): Placed {
		const p: Placed = {
			node: n,
			depth,
			parent,
			leafIdxStart: 0,
			leafCount: 0,
			cx: 0,
			cy: 0,
		};
		nodes.push(p);
		(byDepth[depth] ??= []).push(p);
		if (n.children.length === 0) {
			p.leafIdxStart = leafCounter++;
			p.leafCount = 1;
		} else {
			const start = leafCounter;
			for (const c of n.children) visit(c, depth + 1, p);
			p.leafIdxStart = start;
			p.leafCount = leafCounter - start;
		}
		return p;
	}
	visit(root, 0, null);

	const totalLeaves = Math.max(1, leafCounter);
	const depthCount = byDepth.length;

	// Budget: horizontal per-leaf slot = frameCx / totalLeaves; vertical
	// per-depth slot = frameCy / depthCount. Boxes take ~80% of the slot
	// to leave room for the connector lines.
	const slotW = frameCx / totalLeaves;
	const slotH = frameCy / depthCount;
	const boxW = Math.max(1, Math.round(slotW * 0.85));
	const boxH = Math.max(1, Math.round(slotH * 0.6));

	// Compute centers: leaves first, then propagate upward.
	for (const p of nodes) {
		// Provisional — overwritten for non-leaves below.
		p.cx = Math.round(slotW * (p.leafIdxStart + p.leafCount / 2));
		p.cy = Math.round(slotH * (p.depth + 0.5));
	}
	// Recompute non-leaves bottom-up. Since nodes was filled by a DFS,
	// children always appear after their parents in the array; walking
	// it in reverse visits leaves → internal nodes in the right order.
	for (let i = nodes.length - 1; i >= 0; i--) {
		const p = nodes[i];
		if (p.node.children.length === 0) continue;
		// Centre of children's centre-span.
		let minCx = Infinity;
		let maxCx = -Infinity;
		for (const n of nodes) {
			if (n.parent === p) {
				minCx = Math.min(minCx, n.cx);
				maxCx = Math.max(maxCx, n.cx);
			}
		}
		if (minCx !== Infinity) {
			p.cx = Math.round((minCx + maxCx) / 2);
		}
	}

	const out: ShapeLike[] = [];
	// Emit connector lines first so boxes paint on top.
	for (const p of nodes) {
		if (!p.parent) continue;
		const fromX = p.parent.cx;
		const fromY = p.parent.cy + Math.round(boxH / 2);
		const toX = p.cx;
		const toY = p.cy - Math.round(boxH / 2);
		out.push(makeConnector(fromX, fromY, toX, toY, strokeHex));
	}
	for (const p of nodes) {
		const x = p.cx - Math.round(boxW / 2);
		const y = p.cy - Math.round(boxH / 2);
		out.push(makeBox(x, y, boxW, boxH, p.node.text, accentHex));
	}
	return out;
}

// Factory for a rectangle with centred text. Uses prstGeom="rect" so
// the standard renderer draws it as an ordinary auto-shape, and leaves
// effects/hyperlinks/etc. null. Text runs inherit the master default
// styles through the normal chain since we don't specify sizes.
function makeBox(
	x: number,
	y: number,
	cx: number,
	cy: number,
	text: string,
	accentHex: string,
): ShapeLike {
	return {
		kind: 'shape',
		x,
		y,
		cx,
		cy,
		fill: { kind: 'solid', colorHex: accentHex, alpha: null },
		line: {
			widthEmu: 9525, // ~1px
			fill: { kind: 'solid', colorHex: "FFFFFF", alpha: null },
			dash: null,
			cap: null,
			join: null,
			cmpd: null,
			headEnd: null,
			tailEnd: null,
		},
		paragraphs: text
			? [{
				level: 0,
				style: {
					align: 'ctr',
					marL: null,
					indent: null,
					bullet: { kind: 'none' },
					lineSpacing: null,
					spaceBefore: null,
					spaceAfter: null,
					rtl: null,
					tabs: null,
				},
				runs: [{
					kind: 'text',
					text,
					style: {
						sizeHundredths: null,
						bold: null,
						italic: null,
						colorHex: "FFFFFF",
						colorSchemeSlot: null,
						colorMods: null,
						fontFamily: null,
						underline: null,
						strike: null,
						baseline: null,
						letterSpacingHundredths: null,
						kernHundredths: null,
						alpha: null,
						hyperlinkRId: null,
						lang: null,
					},
				}],
			}]
			: [],
		custGeom: null,
		presetGeom: { name: "rect", avLst: new Map() },
		effects: null,
		name: null,
		title: null,
		alt: null,
		hyperlinkRId: null,
		rotation60k: 0,
		flipH: false,
		flipV: false,
		phType: null,
		phIdx: null,
		isTextBox: false,
		bodyPr: null,
	};
}

// Synthesise a thin connector as a zero-fill shape with a custGeom
// describing a line from (fromX, fromY) to (toX, toY). Using custGeom
// (rather than prstGeom="line") keeps the renderer agnostic — the same
// path rendering that handles arbitrary shapes handles this one.
function makeConnector(
	fromX: number,
	fromY: number,
	toX: number,
	toY: number,
	strokeHex: string,
): ShapeLike {
	const x = Math.min(fromX, toX);
	const y = Math.min(fromY, toY);
	const cx = Math.max(1, Math.abs(toX - fromX));
	const cy = Math.max(1, Math.abs(toY - fromY));
	// Local path in a 100000 × 100000 grid so integer coords don't lose
	// resolution on small boxes.
	const lx = (v: number) => Math.round(((v - x) / cx) * 100000);
	const ly = (v: number) => Math.round(((v - y) / cy) * 100000);
	const d = `M ${lx(fromX)} ${ly(fromY)} L ${lx(toX)} ${ly(toY)}`;
	return {
		kind: 'shape',
		x,
		y,
		cx,
		cy,
		fill: { kind: 'none' },
		line: {
			widthEmu: 6350, // ~0.67px — thinner than the box stroke
			fill: { kind: 'solid', colorHex: strokeHex, alpha: null },
			dash: null,
			cap: null,
			join: null,
			cmpd: null,
			headEnd: null,
			tailEnd: null,
		},
		paragraphs: [],
		custGeom: {
			pathW: 100000,
			pathH: 100000,
			d,
			closed: false,
			fillMode: 'none',
		},
		presetGeom: null,
		effects: null,
		name: null,
		title: null,
		alt: null,
		hyperlinkRId: null,
		rotation60k: 0,
		flipH: false,
		flipV: false,
		phType: null,
		phIdx: null,
		isTextBox: false,
		bodyPr: null,
	};
}
