import type { Shape } from '../presentation-parser';
import type { BodyProperties } from '../text-style';
import type { ShapeEffects, OuterShadow, InnerShadow, Glow, SoftEdge, Reflection, Blur } from '../effects';
import { emuToPx, positionStyle, transformStyle, SVG_NS } from './geom';
import { renderParagraph, type AutoNumState, type BodyTextContext, type FieldContext } from './text';
import { solidColorFromFill, fillToCssBackground } from './fill-utils';
import { presetToSvgPath } from '../preset-geom';
import { wrapInHyperlink } from './hyperlink';

// PowerPoint's default text-frame insets (EMU). Match the values PowerPoint
// uses when <a:bodyPr> omits lIns/tIns/rIns/bIns: 0.1" horizontal, 0.05"
// vertical.
const DEFAULT_L_INS_EMU = 91440;
const DEFAULT_R_INS_EMU = 91440;
const DEFAULT_T_INS_EMU = 45720;
const DEFAULT_B_INS_EMU = 45720;

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
	const isDegenerateLine = shape.cx === 0 || shape.cy === 0;

	if (shape.custGeom && !isDegenerateLine) {
		el.appendChild(renderCustGeomSvg(shape));
	} else if (shape.presetGeom && !isDegenerateLine) {
		const svg = renderPresetGeomSvg(shape);
		if (svg) {
			el.appendChild(svg);
		} else {
			// Unknown preset — fall back to the plain-box look.
			applyBoxFill(el, shape, embedUrls);
		}
	} else {
		applyBoxFill(el, shape, embedUrls);
	}

	if (shape.effects) applyEffects(el, shape.effects);

	if (shape.paragraphs.length > 0) {
		el.appendChild(renderTextBody(shape, hyperlinkUrls, fieldCtx));
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
		wrap.appendChild(renderParagraph(p, autoNumState, hyperlinkUrls, fieldCtx, bodyCtx));
	}
	return wrap;
}

// Emit the fill / border as CSS on the shape's <div> (legacy path for shapes
// with no custGeom or preset, or with an unrecognised preset).
function applyBoxFill(el: HTMLElement, shape: Shape, embedUrls: Map<string, string>): void {
	const bg = fillToCssBackground(shape.fill, embedUrls);
	if (bg) el.style.background = bg;
	// TODO: support gradient / blip / pattern line fills; currently only
	// the solid color path is honoured for strokes.
	const lineColor = solidColorFromFill(shape.line?.fill ?? null);
	if (shape.line && lineColor) {
		const widthPx = shape.line.widthEmu != null ? Math.max(emuToPx(shape.line.widthEmu), 0.5) : 1;
		const isHLine = shape.cy === 0;
		const isVLine = shape.cx === 0;
		if (isHLine || isVLine) {
			el.style.background = lineColor;
			if (isHLine) el.style.height = `${widthPx}px`;
			if (isVLine) el.style.width = `${widthPx}px`;
		} else {
			el.style.border = `${widthPx}px solid ${lineColor}`;
		}
	}
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

function withAlpha(hex: string, alpha: number | null): string {
	if (alpha == null || alpha >= 1) return hex;
	const m = /^#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$/.exec(hex);
	if (!m) return hex;
	const r = parseInt(m[1], 16);
	const g = parseInt(m[2], 16);
	const b = parseInt(m[3], 16);
	return `rgba(${r}, ${g}, ${b}, ${Math.max(0, Math.min(1, alpha))})`;
}

function outerShadowToBoxShadow(s: OuterShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlpha(s.colorHex, s.alpha)}`;
}

function innerShadowToBoxShadow(s: InnerShadow): string {
	const { dx, dy } = dirDistToOffsetPx(s.dir60k, s.distEmu);
	const blurPx = emuToPx(s.blurRadEmu);
	return `inset ${dx.toFixed(2)}px ${dy.toFixed(2)}px ${blurPx.toFixed(2)}px ${withAlpha(s.colorHex, s.alpha)}`;
}

function glowToBoxShadow(g: Glow): string {
	const radPx = emuToPx(g.radEmu);
	return `0 0 ${radPx.toFixed(2)}px ${radPx.toFixed(2)}px ${withAlpha(g.colorHex, g.alpha)}`;
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
function buildShapeSvg(shape: Shape, vbW: number, vbH: number, d: string, closed: boolean): SVGSVGElement {
	const svg = document.createElementNS(SVG_NS, "svg");
	svg.setAttribute("width", "100%");
	svg.setAttribute("height", "100%");
	svg.setAttribute("viewBox", `0 0 ${vbW} ${vbH}`);
	svg.setAttribute("preserveAspectRatio", "none");
	Object.assign(svg.style, { position: "absolute", left: "0", top: "0", width: "100%", height: "100%" });
	const path = document.createElementNS(SVG_NS, "path");
	path.setAttribute("d", d);
	path.setAttribute("fill", shape.fill?.kind === 'solid' ? shape.fill.colorHex : "none");
	const strokeColor = solidColorFromFill(shape.line?.fill ?? null);
	if (strokeColor) {
		path.setAttribute("stroke", strokeColor);
		const widthPx = shape.line!.widthEmu != null ? Math.max(emuToPx(shape.line!.widthEmu), 0.5) : 1;
		path.setAttribute("stroke-width", String(widthPx));
		path.setAttribute("vector-effect", "non-scaling-stroke");
	} else if (!closed && shape.fill?.kind !== 'solid') {
		path.setAttribute("stroke", "#000");
		path.setAttribute("stroke-width", "1");
		path.setAttribute("vector-effect", "non-scaling-stroke");
	}
	svg.appendChild(path);
	return svg;
}

export function renderCustGeomSvg(shape: Shape): SVGSVGElement {
	const vbW = shape.custGeom!.pathW || Math.max(shape.cx, 1);
	const vbH = shape.custGeom!.pathH || Math.max(shape.cy, 1);
	return buildShapeSvg(shape, vbW, vbH, shape.custGeom!.d, shape.custGeom!.closed);
}

export function renderPresetGeomSvg(shape: Shape): SVGSVGElement | null {
	const pg = shape.presetGeom!;
	const vbW = Math.max(shape.cx, 1);
	const vbH = Math.max(shape.cy, 1);
	const d = presetToSvgPath(pg.name, vbW, vbH, pg.avLst);
	if (d == null) return null;
	const closed = pg.name !== "line" && pg.name !== "straightConnector1";
	return buildShapeSvg(shape, vbW, vbH, d, closed);
}
