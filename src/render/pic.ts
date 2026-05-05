import type { PicShape } from '../presentation-parser';
import { positionStyle, transformStyle, SVG_NS } from './geom';
import { wrapInHyperlink } from './hyperlink';
import {
	applyCropOverlay,
	applyTileMetrics,
	createInlineFilter,
	nextSvgFilterId,
	supportsSvgFilters,
	hexTo01,
} from './fill-utils';

// Per-mille denominator used throughout OOXML blip-effect attributes
// (srcRect/alphaModFix/lum/biLevel/tile). 100000 = 100%.
const PERMILLE = 100000;

export function renderPic(pic: PicShape, cls: string, hyperlinkUrls: Map<string, string>): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	Object.assign(wrap.style, transformStyle(pic.rotation60k, pic.flipH, pic.flipV));

	// Stable data-* hooks for the conformance harness / DOM introspection.
	wrap.setAttribute("data-kind", "pic");
	if (pic.phType) {
		wrap.setAttribute("data-placeholder-type", pic.phType);
		if (pic.phIdx != null) wrap.setAttribute("data-placeholder-idx", pic.phIdx);
	}

	if (!pic.src) return wrap;

	const opacity = pic.alphaPermille != null ? pic.alphaPermille / PERMILLE : null;

	if (!pic.stretch) {
		// Tile fill — render as a repeating background on the container.
		// Precise tile sizing/offset requires the source image's natural px
		// dimensions, which we can only learn asynchronously; `applyTileMetrics`
		// patches backgroundSize/Position in place after an Image() probe.
		wrap.style.backgroundImage = `url(${cssUrl(pic.src)})`;
		wrap.style.backgroundRepeat = "repeat";
		wrap.style.backgroundSize = "auto";
		if (pic.tile) {
			applyTileMetrics(wrap, pic.src, pic.tile);
		}
		// Build filters *against the tile wrapper* so any inline <svg> filter
		// defs land alongside the background div.
		const filter = buildFilterString(wrap, pic);
		if (filter) wrap.style.filter = filter;
		if (opacity != null) wrap.style.opacity = String(opacity);
		applyDuotone(wrap, wrap, pic);
		return wrap;
	}

	// Stretch (default) path — emit an <img> sized to fill the wrapper,
	// optionally inside a crop wrapper when <a:srcRect> is present. The
	// crop math is shared with shape.ts via applyCropOverlay so both
	// code paths honour <a:srcRect> identically.
	const img = applyCropOverlay(wrap, pic.src, pic.srcRectPermille);
	// The DOCX-side render guidance warns against putting document-derived
	// strings into innerHTML/CSS; `alt`/`title` via setAttribute is safe
	// (the browser HTML-encodes attribute values).
	// PicShape.alt already falls back descr → name → "" during parse, so
	// it's the right value for <img alt=>. Only emit `title` when the
	// source set an explicit `title` attribute on the picture.
	if (pic.alt) img.setAttribute("alt", pic.alt);
	if (pic.title) img.setAttribute("title", pic.title);

	// Filter defs live inside `wrap` so they share the DOM subtree; the
	// filter URL is applied to the <img> itself (which actually carries
	// the pixel data the filter consumes).
	const filter = buildFilterString(wrap, pic);
	if (filter) img.style.filter = filter;
	if (opacity != null) img.style.opacity = String(opacity);

	applyDuotone(wrap, img, pic);

	// Hyperlinked image: wrap and hoist positioning onto the anchor so the
	// click area matches the pic box.
	if (pic.hyperlinkRId && hyperlinkUrls.has(pic.hyperlinkRId)) {
		const anchorStyles = {
			position: wrap.style.position,
			left: wrap.style.left,
			top: wrap.style.top,
			width: wrap.style.width,
			height: wrap.style.height,
		};
		wrap.style.position = "relative";
		wrap.style.left = "0";
		wrap.style.top = "0";
		wrap.style.width = "100%";
		wrap.style.height = "100%";
		const wrapped = wrapInHyperlink(wrap, pic.hyperlinkRId, hyperlinkUrls);
		if (wrapped !== wrap) {
			Object.assign(wrapped.style, anchorStyles, { display: "block" });
		}
		return wrapped;
	}
	return wrap;
}

// Compose every per-pic filter into a single CSS `filter` string. For
// effects with a clean CSS analogue (grayscale / brightness / contrast)
// we stick with the native function — browsers implement them with the
// same feColorMatrix machinery under the hood. biLevel, which the CSS
// approximation rendered as near-pure black/white with a tilt toward
// "brightness+contrast(1000)", is upgraded to an inline SVG <filter> so
// the threshold actually lands where PowerPoint puts it.
function buildFilterString(defsHost: HTMLElement, pic: PicShape): string {
	const parts: string[] = [];
	if (pic.grayscale) parts.push("grayscale(1)");
	if (pic.lumBrightPermille != null) {
		parts.push(`brightness(${1 + pic.lumBrightPermille / PERMILLE})`);
	}
	if (pic.lumContrastPermille != null) {
		parts.push(`contrast(${1 + pic.lumContrastPermille / PERMILLE})`);
	}
	if (pic.biLevelPermille != null) {
		const biLevelFilter = buildBiLevelFilter(defsHost, pic.biLevelPermille);
		if (biLevelFilter) {
			parts.push(biLevelFilter);
		} else {
			// Fallback: the grayscale+brightness+contrast(1000) trick.
			const threshFrac = pic.biLevelPermille / PERMILLE;
			const bright = 1 + (0.5 - threshFrac);
			parts.push(`grayscale(1)`);
			parts.push(`brightness(${bright})`);
			parts.push(`contrast(1000)`);
		}
	}
	return parts.join(" ");
}

// Apply a duotone effect. Prefers an SVG <filter> (feColorMatrix to extract
// luminance, feComponentTransfer to map 0..1 luminance to a ramp between
// the two colours). Falls back to the legacy background + mix-blend-mode
// overlay when SVGFilterElement isn't available — the legacy approach is
// visibly wrong (it multiplies, it doesn't map luminance) but keeps some
// visual signal in environments where the filter won't paint.
function applyDuotone(defsHost: HTMLElement, target: HTMLElement, pic: PicShape): void {
	if (!pic.duotone) return;
	if (supportsSvgFilters()) {
		const filterUrl = buildDuotoneFilter(defsHost, pic.duotone);
		if (filterUrl) {
			// Chain with any filter already on the target (grayscale / biLevel).
			const existing = target.style.filter;
			target.style.filter = existing ? `${existing} ${filterUrl}` : filterUrl;
			return;
		}
	}
	// Legacy fallback — mix-blend-mode overlay on the wrapper.
	defsHost.style.backgroundColor = pic.duotone[0];
	const overlay = document.createElement("div");
	overlay.style.position = "absolute";
	overlay.style.inset = "0";
	overlay.style.backgroundColor = pic.duotone[1];
	overlay.style.mixBlendMode = "multiply";
	overlay.style.pointerEvents = "none";
	defsHost.appendChild(overlay);
}

// Build a `<filter>` that maps the image's luminance onto a two-colour
// ramp. Returns `url(#id)` or null when the filter couldn't be constructed
// (only if supportsSvgFilters() returns false).
function buildDuotoneFilter(defsHost: HTMLElement, duotone: [string, string]): string | null {
	if (!supportsSvgFilters()) return null;
	const [dark, light] = duotone;
	const [dr, dg, db] = hexTo01(dark);
	const [lr, lg, lb] = hexTo01(light);
	const id = nextSvgFilterId('pptx-duotone');
	const { svg, filter } = createInlineFilter(id);

	// Step 1 — collapse to luminance. Rec. 601 weights (0.299/0.587/0.114)
	// in all three channels give a gray-equivalent RGBA where R=G=B=Y.
	const cm = document.createElementNS(SVG_NS, 'feColorMatrix');
	cm.setAttribute('type', 'matrix');
	cm.setAttribute('values', [
		'0.299 0.587 0.114 0 0',
		'0.299 0.587 0.114 0 0',
		'0.299 0.587 0.114 0 0',
		'0     0     0     1 0',
	].join('\n'));
	cm.setAttribute('result', 'luma');
	filter.appendChild(cm);

	// Step 2 — remap each channel's 0..1 luminance onto the dark→light ramp
	// via feComponentTransfer/feFuncX with `type="table"` (a two-entry
	// table does a linear interpolation between the two endpoints).
	const ct = document.createElementNS(SVG_NS, 'feComponentTransfer');
	ct.setAttribute('in', 'luma');
	const fr = document.createElementNS(SVG_NS, 'feFuncR');
	fr.setAttribute('type', 'table');
	fr.setAttribute('tableValues', `${dr.toFixed(4)} ${lr.toFixed(4)}`);
	const fg = document.createElementNS(SVG_NS, 'feFuncG');
	fg.setAttribute('type', 'table');
	fg.setAttribute('tableValues', `${dg.toFixed(4)} ${lg.toFixed(4)}`);
	const fb = document.createElementNS(SVG_NS, 'feFuncB');
	fb.setAttribute('type', 'table');
	fb.setAttribute('tableValues', `${db.toFixed(4)} ${lb.toFixed(4)}`);
	ct.appendChild(fr);
	ct.appendChild(fg);
	ct.appendChild(fb);
	filter.appendChild(ct);

	defsHost.appendChild(svg);
	return `url(#${id})`;
}

// Build a `<filter>` that thresholds the image's luminance at `permille`
// (OOXML per-mille = 100000 ⇒ 100%). The feFuncX step uses `type="linear"`
// with a steep slope/intercept so everything above the threshold snaps to
// 1.0 and everything below snaps to 0.0 after the implicit [0,1] clamp.
// Returns null when SVG filters aren't supported so the caller can fall
// back to the grayscale + brightness + contrast(1000) approximation.
function buildBiLevelFilter(defsHost: HTMLElement, permille: number): string | null {
	if (!supportsSvgFilters()) return null;
	const threshold = Math.max(0, Math.min(1, permille / PERMILLE));
	const id = nextSvgFilterId('pptx-bilevel');
	const { svg, filter } = createInlineFilter(id);

	// Luminance matrix: collapse RGB to a single gray channel.
	const cm = document.createElementNS(SVG_NS, 'feColorMatrix');
	cm.setAttribute('type', 'matrix');
	cm.setAttribute('values', [
		'0.299 0.587 0.114 0 0',
		'0.299 0.587 0.114 0 0',
		'0.299 0.587 0.114 0 0',
		'0     0     0     1 0',
	].join('\n'));
	cm.setAttribute('result', 'luma');
	filter.appendChild(cm);

	// Threshold via feComponentTransfer type="linear". y = slope*x + intercept,
	// then clamped to [0,1] by the filter pipeline. A slope of 1e6 puts any
	// input above the threshold firmly above 1.0 (→ clamped to 1) and any
	// input below firmly below 0 (→ clamped to 0). The intercept is set so
	// that x == threshold produces exactly y == 0, i.e. intercept = -slope*threshold.
	const slope = 1e6;
	const intercept = -slope * threshold;
	const ct = document.createElementNS(SVG_NS, 'feComponentTransfer');
	ct.setAttribute('in', 'luma');
	for (const tag of ['feFuncR', 'feFuncG', 'feFuncB']) {
		const f = document.createElementNS(SVG_NS, tag);
		f.setAttribute('type', 'linear');
		f.setAttribute('slope', String(slope));
		f.setAttribute('intercept', intercept.toExponential(6));
		ct.appendChild(f);
	}
	filter.appendChild(ct);

	defsHost.appendChild(svg);
	return `url(#${id})`;
}

function cssUrl(src: string): string {
	const escaped = src.replace(/(["\\])/g, "\\$1").replace(/\n/g, "\\A ");
	return `"${escaped}"`;
}
