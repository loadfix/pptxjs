import type { Shape } from '../presentation-parser';
import type { BodyProperties } from '../text-style';
import type { ShapeEffects, OuterShadow, InnerShadow, Glow, SoftEdge, Reflection, Blur } from '../effects';
import type { LineStyle, LineEnd } from '../fill';
import { emuToPx, positionStyle, transformStyle, SVG_NS } from './geom';
import { renderParagraph, type AutoNumState, type BodyTextContext, type FieldContext } from './text';
import {
	approxSolidColorFromFill,
	fillToCssBackground,
	fillToSvgPaint,
	applyCropOverlay,
	lineDashToCss,
	svgDashArray,
} from './fill-utils';
import { withAlphaHex } from '../color-math';
import { presetToSvgPath } from '../preset-geom';
import { wrapInHyperlink } from './hyperlink';

// PowerPoint's default text-frame insets (EMU). Match the values PowerPoint
// uses when <a:bodyPr> omits lIns/tIns/rIns/bIns: 0.1" horizontal, 0.05"
// vertical.
const DEFAULT_L_INS_EMU = 91440;
const DEFAULT_R_INS_EMU = 91440;
const DEFAULT_T_INS_EMU = 45720;
const DEFAULT_B_INS_EMU = 45720;

// Monotonically increasing id used to namespace marker defs so multiple
// shapes on the same document don't collide on `url(#head-foo)`.
let markerIdCounter = 0;

export function renderShape(
	shape: Shape,
	cls: string,
	embedUrls: Map<string, string>,
	hyperlinkUrls: Map<string, string>,
	fieldCtx?: FieldContext,
): HTMLElement {
	const el = document.createElement("div");
	el.className = `${cls}-shape`;
	Object.assign(el.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));
	Object.assign(el.style, transformStyle(shape.rotation60k, shape.flipH, shape.flipV));

	// Accessibility: prefer `title` (human-authored title text) for aria-label;
	// fall back to `alt` (from <p:cNvPr descr=>). Skip entirely when neither is
	// set so we don't emit empty ARIA attributes that hurt screen-reader UX.
	const ariaLabel = shape.title || shape.alt;
	if (ariaLabel) {
		el.setAttribute("aria-label", ariaLabel);
		el.setAttribute("role", "figure");
	}

	// A degenerate "line" (zero cx or cy) can't render a path in zero area,
	// so fall back to the background-band treatment even if custGeom is set.
	// Exception: when the line carries an arrow head, we render a tiny inline
	// SVG instead so the <marker> actually has something to attach to.
	const isDegenerateLine = shape.cx === 0 || shape.cy === 0;
	const hasArrowHead = lineHasArrowHead(shape.line);

	if (shape.custGeom && !isDegenerateLine) {
		el.appendChild(renderCustGeomSvg(shape, embedUrls));
	} else if (shape.presetGeom && !isDegenerateLine) {
		const svg = renderPresetGeomSvg(shape, embedUrls);
		if (svg) {
			el.appendChild(svg);
		} else {
			// Unknown preset — fall back to the plain-box look.
			applyBoxFill(el, shape, embedUrls);
		}
	} else if (isDegenerateLine && hasArrowHead && shape.line?.fill && shape.line.fill.kind !== 'none') {
		// Arrow-head-bearing connector with a degenerate dimension — render
		// as an oversized inline SVG so the marker has space to draw.
		renderDegenerateLineSvg(el, shape, embedUrls);
	} else {
		applyBoxFill(el, shape, embedUrls);
	}

	if (shape.effects) applyEffects(el, shape.effects);

	if (shape.paragraphs.length > 0) {
		el.appendChild(renderTextBody(shape, hyperlinkUrls, embedUrls, fieldCtx));
	}
	// Whole-shape click-action: wrap the positioned box in an <a>. The anchor
	// inherits the shape's position so the hit area is the shape itself.
	if (shape.hyperlinkRId && hyperlinkUrls.has(shape.hyperlinkRId)) {
		// Move positioning from the shape to the anchor so the <a> occupies the
		// shape's click area — otherwise the anchor collapses to zero size.
		const anchorStyles = {
			position: el.style.position,
			left: el.style.left,
			top: el.style.top,
			width: el.style.width,
			height: el.style.height,
		};
		el.style.position = "relative";
		el.style.left = "0";
		el.style.top = "0";
		el.style.width = "100%";
		el.style.height = "100%";
		const wrapped = wrapInHyperlink(el, shape.hyperlinkRId, hyperlinkUrls);
		if (wrapped !== el) {
			Object.assign(wrapped.style, anchorStyles, { display: "block" });
		}
		return wrapped;
	}
	return el;
}

// Build the text-body wrapper <div> for a shape: applies <a:bodyPr> insets,
// vertical anchor, wrap, column count, writing-mode, then renders each
// paragraph inside. Autofit's fontScale / lnSpcReduction ride along via
// `BodyTextContext` so every run picks them up uniformly. Returns a single
// element positioned absolutely over the shape (so it stacks above any
// SVG-drawn geometry).
function renderTextBody(
	shape: Shape,
	hyperlinkUrls: Map<string, string>,
	embedUrls: Map<string, string>,
	fieldCtx?: FieldContext,
): HTMLElement {
	const bp: BodyProperties | null = shape.bodyPr;

	const wrap = document.createElement("div");
	// Fill the shape box; absolute so it sits above background SVG/preset.
	Object.assign(wrap.style, {
		position: "absolute",
		left: "0",
		top: "0",
		width: "100%",
		height: "100%",
		boxSizing: "border-box",
	} as Partial<CSSStyleDeclaration>);

	// Insets → padding. Apply PowerPoint defaults when an attribute is unset.
	const lIns = bp?.lInsEmu ?? DEFAULT_L_INS_EMU;
	const tIns = bp?.tInsEmu ?? DEFAULT_T_INS_EMU;
	const rIns = bp?.rInsEmu ?? DEFAULT_R_INS_EMU;
	const bIns = bp?.bInsEmu ?? DEFAULT_B_INS_EMU;
	wrap.style.padding =
		`${emuToPx(tIns)}px ${emuToPx(rIns)}px ${emuToPx(bIns)}px ${emuToPx(lIns)}px`;

	// Vertical anchor → flexbox justify-content on a column flex container.
	// anchor=t (or unset) pins to the top; ctr centers; b pins to bottom.
	// 'just'/'dist' are approximated as top — distributing paragraphs evenly
	// is not something CSS does well without per-line hooks.
	wrap.style.display = "flex";
	wrap.style.flexDirection = "column";
	switch (bp?.anchor) {
		case 'ctr': wrap.style.justifyContent = "center"; break;
		case 'b': wrap.style.justifyContent = "flex-end"; break;
		default: wrap.style.justifyContent = "flex-start"; break;
	}

	// wrap="none" disables wrapping; default (square) allows normal wrapping.
	if (bp?.wrap === 'none') {
		wrap.style.whiteSpace = "pre";
	}

	// Multi-column text frames — <a:bodyPr numCol="2" spcCol="N">.
	if (bp?.numCol != null && bp.numCol > 1) {
		wrap.style.columnCount = String(bp.numCol);
		if (bp.spcColEmu != null) {
			wrap.style.columnGap = `${emuToPx(bp.spcColEmu)}px`;
		}
	}

	// Vertical writing-mode. `vert270` wants bottom-up vertical text which CSS
	// expresses as sideways-lr; older browsers ignore that keyword and fall
	// back to vertical-rl (top-down), which is still readable. mongolianVert
	// and the wordArt* variants are uncommon and left as TODO.
	switch (bp?.vert) {
		case 'vert':
		case 'eaVert':
			wrap.style.writingMode = "vertical-rl";
			break;
		case 'vert270':
			wrap.style.writingMode = "sideways-lr";
			break;
		// TODO: mongolianVert, wordArtVert, wordArtVertRtl.
	}

	// Autofit — normAutofit shrinks every run's font-size and (optionally)
	// line-height uniformly. Threaded into renderParagraph so per-run pt
	// values multiply without CSS custom-property gymnastics.
	let bodyCtx: BodyTextContext | undefined;
	if (bp?.autofit?.kind === 'normAutofit') {
		bodyCtx = {
			fontScale: bp.autofit.fontScale,
			lnSpcReduction: bp.autofit.lnSpcReduction,
		};
	}

	const autoNumState: AutoNumState = new Map();
	for (const p of shape.paragraphs) {
		wrap.appendChild(renderParagraph(p, autoNumState, hyperlinkUrls, fieldCtx, bodyCtx, embedUrls));
	}
	return wrap;
}

// Emit the fill / border as CSS on the shape's <div> (legacy path for shapes
// with no custGeom or preset, or with an unrecognised preset).
function applyBoxFill(el: HTMLElement, shape: Shape, embedUrls: Map<string, string>): void {
	const bg = fillToCssBackground(shape.fill, embedUrls);
	if (bg) {
		if (bg.kind === 'css') {
			el.style.background = bg.value;
		} else {
			// blipFill — honour <a:srcRect> via an overlay <img> so crops
			// render correctly. Plain `background-image` can't crop to a
			// sub-rectangle; the overlay wrapper is overflow:hidden and the
			// <img> is positioned/scaled to expose only the visible crop.
			if (bg.tile) {
				// TODO: tile mode still uses background-repeat (same caveat
				// as before — srcRect can't be honoured for tiling in CSS).
				el.style.backgroundImage = `url("${bg.src.replace(/"/g, '\\"')}")`;
				el.style.backgroundRepeat = 'repeat';
				el.style.backgroundPosition = 'top left';
				el.style.backgroundSize = 'auto';
			} else {
				if (!el.style.position) el.style.position = 'absolute';
				applyCropOverlay(el, bg.src, bg.srcRectPermille);
			}
		}
	}
	// Line/stroke. The CSS border path can only carry a flat color; for
	// non-solid line fills we fall back to a best-effort approximation:
	// gradient → first stop, pattern → fg color, blip → a neutral grey
	// (blips can't be painted into a CSS border — TODO: use an SVG overlay
	// when blip strokes on plain rectangles matter enough). Shapes that go
	// through the SVG path instead pick up full gradient/pattern/blip
	// strokes via fillToSvgPaint.
	let lineColor = approxSolidColorFromFill(shape.line?.fill ?? null);
	if (lineColor == null && shape.line?.fill?.kind === 'blip') lineColor = '#808080';
	if (shape.line && lineColor) {
		const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
		const isHLine = shape.cy === 0;
		const isVLine = shape.cx === 0;
		if (isHLine || isVLine) {
			el.style.background = lineColor;
			if (isHLine) el.style.height = `${widthPx}px`;
			if (isVLine) el.style.width = `${widthPx}px`;
		} else {
			// cmpd=dbl maps cleanly to border-style: double. thickThin /
			// thinThick have no CSS analogue — pick the thicker of the two
			// visual weights and fall through to a single border. tri is
			// unsupported and handled the same way.
			const borderStyle = shape.line.cmpd === 'dbl'
				? 'double'
				: lineDashToCss(shape.line.dash);
			// For `double`, CSS needs at least 3px of width to render two
			// visible lines with a gap. Bump narrow double borders up.
			const effectiveWidth = borderStyle === 'double'
				? Math.max(widthPx, 3)
				: widthPx;
			el.style.border = `${effectiveWidth}px ${borderStyle} ${lineColor}`;
		}
	}
}

// True when the line has a non-'none' head or tail end marker. The parser
// defaults `type` to 'none' when the attribute is omitted, so a present-but-
// empty <a:headEnd/> also counts as no-arrow.
function lineHasArrowHead(line: LineStyle | null | undefined): boolean {
	if (!line) return false;
	const h = line.headEnd?.type;
	const t = line.tailEnd?.type;
	return (h != null && h !== 'none') || (t != null && t !== 'none');
}

// Map a parsed <a:effectLst> onto CSS on the outer wrapper div.
function applyEffects(el: HTMLElement, fx: ShapeEffects): void {
	const boxShadows: string[] = [];
	const filters: string[] = [];

	if (fx.outerShadow) boxShadows.push(outerShadowToBoxShadow(fx.outerShadow));
	if (fx.innerShadow) boxShadows.push(innerShadowToBoxShadow(fx.innerShadow));
	if (fx.glow) boxShadows.push(glowToBoxShadow(fx.glow));

	if (fx.blur) filters.push(blurToFilter(fx.blur));
	if (fx.softEdge) filters.push(softEdgeToFilter(fx.softEdge));

	if (boxShadows.length) el.style.boxShadow = boxShadows.join(', ');
	if (filters.length) el.style.filter = filters.join(' ');

	if (fx.reflection) {
		const r = reflectionToWebkitBoxReflect(fx.reflection);
		if (r) (el.style as CSSStyleDeclaration & { webkitBoxReflect?: string }).webkitBoxReflect = r;
	}
}

function dirDistToOffsetPx(dir60k: number, distEmu: number): { dx: number; dy: number } {
	const rad = (dir60k / 60000) * Math.PI / 180;
	const distPx = emuToPx(distEmu);
	return { dx: distPx * Math.cos(rad), dy: distPx * Math.sin(rad) };
}

function outerShadowToBoxShadow(s: OuterShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlphaHex(s.colorHex, s.alpha)}`;
}

function innerShadowToBoxShadow(s: InnerShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `inset ${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlphaHex(s.colorHex, s.alpha)}`;
}

function glowToBoxShadow(g: Glow): string {
	const radPx = emuToPx(g.radEmu);
	return `0 0 ${radPx.toFixed(2)}px ${radPx.toFixed(2)}px ${withAlphaHex(g.colorHex, g.alpha)}`;
}

function blurToFilter(b: Blur): string {
	const px = emuToPx(b.radEmu);
	return `blur(${px.toFixed(2)}px)`;
}

function softEdgeToFilter(s: SoftEdge): string {
	const px = emuToPx(s.radEmu);
	return `blur(${px.toFixed(2)}px)`;
}

function reflectionToWebkitBoxReflect(r: Reflection): string | null {
	const gapPx = emuToPx(r.distEmu);
	const stA = Math.max(0, Math.min(1, r.stAPermille / 100000));
	const endA = Math.max(0, Math.min(1, r.endAPermille / 100000));
	const mask = `linear-gradient(to bottom, rgba(0,0,0,${stA.toFixed(3)}), rgba(0,0,0,${endA.toFixed(3)}))`;
	return `below ${gapPx.toFixed(2)}px ${mask}`;
}

// Shared machinery that turns a (vbW, vbH, d) triple into an <svg> matching
// the shape's fill/stroke styling. Used by both custGeom and presetGeom.
// `fillOverride === 'none'` forces the path's fill to "none" regardless of
// shape.fill — used when <a:path fill="none"> explicitly suppresses fill.
function buildShapeSvg(
	shape: Shape,
	vbW: number,
	vbH: number,
	d: string,
	closed: boolean,
	embedUrls: Map<string, string>,
	fillOverride?: 'none',
): SVGSVGElement {
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	svg.setAttribute("preserveAspectRatio", "none");
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%" });

	// Lazy <defs> — both fill and stroke paint servers (gradients, patterns,
	// blips) drop their elements in here.
	const defs = document.createElementNS(SVG_NS, "defs");

	const path = document.createElementNS(SVG_NS, "path");
	path.setAttribute("d", d);

	// Fill — supports the full union (solid/gradient/pattern/blip) via
	// fillToSvgPaint. `fillOverride === 'none'` takes precedence.
	if (fillOverride === 'none') {
		path.setAttribute("fill", "none");
	} else {
		const fillPaint = fillToSvgPaint(shape.fill, embedUrls, 'pptxjs-fill');
		if (fillPaint) {
			path.setAttribute("fill", fillPaint.paint);
			for (const d of fillPaint.defs) defs.appendChild(d);
		} else {
			path.setAttribute("fill", "none");
		}
	}

	// Stroke — supports the full fill union as well. Gradients/patterns/blips
	// land in <defs> and the stroke attr references them via `url(#id)`.
	const strokePaint = shape.line ? fillToSvgPaint(shape.line.fill, embedUrls, 'pptxjs-stroke') : null;
	if (strokePaint && shape.line) {
		path.setAttribute("stroke", strokePaint.paint);
		for (const d of strokePaint.defs) defs.appendChild(d);
		const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
		path.setAttribute("stroke-width", String(widthPx));
		path.setAttribute("vector-effect", "non-scaling-stroke");
		// Marker colors match the stroke's best-effort solid approximation
		// (markers can't reference paint servers cleanly across renderers).
		const markerColor = approxSolidColorFromFill(shape.line.fill) ?? '#000';
		applyStrokeStyling(svg, path, shape.line, markerColor);
	} else if (!closed && path.getAttribute("fill") === "none") {
		path.setAttribute("stroke", "#000");
		path.setAttribute("stroke-width", "1");
		path.setAttribute("vector-effect", "non-scaling-stroke");
	}

	if (defs.childNodes.length > 0) svg.appendChild(defs);
	svg.appendChild(path);
	return svg;
}

// Emit `stroke-dasharray`, `stroke-linecap`, `stroke-linejoin`, and
// `marker-start`/`marker-end` on `path`, appending any required <marker>
// defs to `svg`. Colour is passed in so the marker fill matches the stroke.
function applyStrokeStyling(svg: SVGSVGElement, path: SVGPathElement, line: LineStyle, strokeColor: string): void {
	const dashArr = svgDashArray(line.dash);
	if (dashArr) path.setAttribute("stroke-dasharray", dashArr);

	if (line.cap) {
		const capMap = { flat: "butt", sq: "square", rnd: "round" } as const;
		path.setAttribute("stroke-linecap", capMap[line.cap]);
	}
	if (line.join) {
		// join values ('round' | 'bevel' | 'miter') already line up with SVG's.
		path.setAttribute("stroke-linejoin", line.join);
	}

	const headMarkerId = line.headEnd && line.headEnd.type !== 'none'
		? addArrowMarker(svg, line.headEnd, strokeColor, 'start')
		: null;
	const tailMarkerId = line.tailEnd && line.tailEnd.type !== 'none'
		? addArrowMarker(svg, line.tailEnd, strokeColor, 'end')
		: null;
	if (headMarkerId) path.setAttribute("marker-start", `url(#${headMarkerId})`);
	if (tailMarkerId) path.setAttribute("marker-end", `url(#${tailMarkerId})`);
}

// Translate w/len size hints onto a scale factor multiplied against stroke
// width. Missing means "medium" per the DrawingML default.
function endSizeFactor(hint: string | null): number {
	switch (hint) {
		case 'sm': return 1.5;
		case 'lg': return 2.5;
		case 'med':
		case null:
		case undefined:
		default:
			return 2;
	}
}

// Add a <marker> def inside <svg><defs>, returning the id so it can be
// referenced by the path. `orient` is 'start' (headEnd) or 'end' (tailEnd);
// the marker path points right (+x) and `orient="auto"` handles the rest.
function addArrowMarker(svg: SVGSVGElement, end: LineEnd, color: string, which: 'start' | 'end'): string | null {
	const wFactor = endSizeFactor(end.w);
	const lFactor = endSizeFactor(end.len);
	// Each marker lives in its own 10x10 viewport; the path is oriented so
	// the tip is at (10, 5) and the base at (0, 0)→(0, 10). Marker units are
	// "strokeWidth" so stroke-width pixels become 1 user unit.
	const vbW = 10;
	const vbH = 10;
	let d: string;
	let markerFill: string | 'none' = color;
	let markerStroke: string | 'none' = 'none';
	switch (end.type) {
		case 'triangle':
			d = `M 0 0 L 10 5 L 0 10 Z`;
			break;
		case 'arrow':
			// Open arrowhead — two strokes meeting at a point.
			d = `M 0 0 L 10 5 L 0 10`;
			markerFill = 'none';
			markerStroke = color;
			break;
		case 'stealth':
			// Concave "stealth" — like triangle but with notched base.
			d = `M 0 0 L 10 5 L 0 10 L 4 5 Z`;
			break;
		case 'diamond':
			d = `M 0 5 L 5 0 L 10 5 L 5 10 Z`;
			break;
		case 'oval':
			// Marker ellipse as a path (circle approximation via 4 arcs).
			d = `M 0 5 A 5 5 0 1 0 10 5 A 5 5 0 1 0 0 5 Z`;
			break;
		case 'none':
		default:
			return null;
	}

	markerIdCounter += 1;
	const id = `pptxjs-arrow-${markerIdCounter}`;

	// Lazy <defs> — reuse if already present so multiple arrows on the same
	// shape share a single <defs> node.
	let defs = svg.querySelector("defs");
	if (!defs) {
		defs = document.createElementNS(SVG_NS, "defs");
		svg.insertBefore(defs, svg.firstChild);
	}

	const marker = document.createElementNS(SVG_NS, "marker");
	marker.setAttribute("id", id);
	marker.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	marker.setAttribute("markerUnits", "strokeWidth");
	// Size scales per w/len hints; len governs along-axis length, w governs
	// perpendicular width.
	marker.setAttribute("markerWidth", String(lFactor * 3));
	marker.setAttribute("markerHeight", String(wFactor * 3));
	marker.setAttribute("refX", which === 'start' ? "0" : "10");
	marker.setAttribute("refY", "5");
	marker.setAttribute("orient", "auto");
	const mp = document.createElementNS(SVG_NS, "path");
	mp.setAttribute("d", d);
	mp.setAttribute("fill", markerFill);
	if (markerStroke !== 'none') {
		mp.setAttribute("stroke", markerStroke);
		mp.setAttribute("stroke-width", "1.5");
		mp.setAttribute("stroke-linecap", "butt");
		mp.setAttribute("stroke-linejoin", "miter");
	}
	marker.appendChild(mp);
	defs.appendChild(marker);
	return id;
}

// Render a degenerate-dimension line (cx=0 or cy=0) as an SVG so arrow-head
// markers have room to draw. Expands the container slightly along the
// vanishing axis so the marker isn't clipped. Uses pixel coordinates in the
// viewBox so marker sizing doesn't get skewed by preserveAspectRatio scaling.
function renderDegenerateLineSvg(el: HTMLElement, shape: Shape, embedUrls: Map<string, string>): void {
	const strokePaint = fillToSvgPaint(shape.line!.fill, embedUrls, 'pptxjs-stroke');
	if (!strokePaint) return;
	const markerColor = approxSolidColorFromFill(shape.line!.fill) ?? '#000';
	const widthPx = shape.line!.widthEmu != null ? Math.max(emuToPx(shape.line!.widthEmu), 0.5) : 1;
	// Rough marker size budget — diameter of the largest marker we'll draw,
	// in pixels. `endSizeFactor` tops out at 2.5x, and markerUnits=strokeWidth
	// means actual size = factor * 3 * strokeWidth. Use a safety margin.
	const tailFactor = Math.max(endSizeFactor(shape.line!.headEnd?.len ?? null), endSizeFactor(shape.line!.tailEnd?.len ?? null));
	const marginPx = Math.ceil(tailFactor * 3 * widthPx) + 2;
	const isHLine = shape.cy === 0;
	const lengthPx = isHLine ? emuToPx(shape.cx) : emuToPx(shape.cy);
	const thicknessPx = widthPx + 2 * marginPx;

	// Override the bounding box so the SVG has some area to render into, and
	// offset it so the centerline aligns with the original box.
	if (isHLine) {
		el.style.height = `${thicknessPx}px`;
		el.style.top = `${emuToPx(shape.y) - marginPx}px`;
	} else {
		el.style.width = `${thicknessPx}px`;
		el.style.left = `${emuToPx(shape.x) - marginPx}px`;
	}

	const vbW = isHLine ? lengthPx : thicknessPx;
	const vbH = isHLine ? thicknessPx : lengthPx;
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	// Preserve aspect ratio so markers render as squares, not stretched.
	svg.setAttribute("preserveAspectRatio", "none");
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%", overflow: "visible" });
	// Paint-server defs (gradient/pattern/blip strokes) go in a lazy <defs>.
	if (strokePaint.defs.length > 0) {
		const defs = document.createElementNS(SVG_NS, "defs");
		for (const d of strokePaint.defs) defs.appendChild(d);
		svg.appendChild(defs);
	}
	const path = document.createElementNS(SVG_NS, "path");
	// Draw the line through the middle of the thickness axis.
	const mid = thicknessPx / 2;
	path.setAttribute(
		"d",
		isHLine ? `M 0 ${mid} L ${lengthPx} ${mid}` : `M ${mid} 0 L ${mid} ${lengthPx}`,
	);
	path.setAttribute("fill", "none");
	path.setAttribute("stroke", strokePaint.paint);
	path.setAttribute("stroke-width", String(widthPx));
	applyStrokeStyling(svg, path, shape.line!, markerColor);
	svg.appendChild(path);
	el.appendChild(svg);
}

export function renderCustGeomSvg(shape: Shape, embedUrls: Map<string, string>): SVGSVGElement {
	const vbW = shape.custGeom!.pathW || Math.max(shape.cx, 1);
	const vbH = shape.custGeom!.pathH || Math.max(shape.cy, 1);
	const fillOverride = shape.custGeom!.fillMode === 'none' ? 'none' : undefined;
	return buildShapeSvg(shape, vbW, vbH, shape.custGeom!.d, shape.custGeom!.closed, embedUrls, fillOverride);
}

export function renderPresetGeomSvg(shape: Shape, embedUrls: Map<string, string>): SVGSVGElement | null {
	const pg = shape.presetGeom!;
	const vbW = Math.max(shape.cx, 1);
	const vbH = Math.max(shape.cy, 1);
	const d = presetToSvgPath(pg.name, vbW, vbH, pg.avLst);
	if (d == null) return null;
	const closed = pg.name !== "line" && pg.name !== "straightConnector1";
	return buildShapeSvg(shape, vbW, vbH, d, closed, embedUrls);
}
