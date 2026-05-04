import type { PicShape } from '../presentation-parser';
import { positionStyle } from './geom';

// Per-mille denominator used throughout OOXML blip-effect attributes
// (srcRect/alphaModFix/lum/biLevel/tile). 100000 = 100%.
const PERMILLE = 100000;

export function renderPic(pic: PicShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));

	if (!pic.src) return wrap;

	// Build the CSS filter list once — applies to both stretch and tile
	// rendering paths.
	const filter = buildFilterString(pic);
	const opacity = pic.alphaPermille != null ? pic.alphaPermille / PERMILLE : null;

	if (!pic.stretch) {
		// Tile fill — render as a repeating background on the container.
		// Offsets (tx/ty) and scales (sx/sy) are in per-mille of the tile
		// source dimensions. We can't resolve the tile source's native
		// pixel size without loading the image, so emit the simplest
		// faithful approximation: repeating at native size.
		wrap.style.backgroundImage = `url(${cssUrl(pic.src)})`;
		wrap.style.backgroundRepeat = "repeat";
		wrap.style.backgroundSize = "auto";
		if (pic.tile) {
			// Translate-offset (per-mille of the tile). We don't know tile
			// size in px, but a percentage of the container is a reasonable
			// visual approximation and avoids JS-side image probing.
			wrap.style.backgroundPosition = `${pic.tile.txPermille / 1000}% ${pic.tile.tyPermille / 1000}%`;
		}
		if (filter) wrap.style.filter = filter;
		if (opacity != null) wrap.style.opacity = String(opacity);
		applyDuotoneFallback(wrap, pic);
		return wrap;
	}

	// Stretch (default) path — emit an <img> sized to fill the wrapper,
	// optionally inside a crop wrapper when <a:srcRect> is present.
	const img = document.createElement("img");
	img.src = pic.src;
	if (pic.alt) img.setAttribute("alt", pic.alt);

	if (pic.srcRectPermille) {
		// Crop via overflow-hidden wrapper + scaled/offset <img>.
		const { l, t, r, b } = pic.srcRectPermille;
		const cropWFrac = 1 - (l + r) / PERMILLE;
		const cropHFrac = 1 - (t + b) / PERMILLE;
		// Guard against pathological rects that would divide by zero or go
		// negative — fall back to uncropped rendering.
		if (cropWFrac > 0 && cropHFrac > 0) {
			wrap.style.overflow = "hidden";
			img.style.position = "absolute";
			img.style.width = `${100 / cropWFrac}%`;
			img.style.height = `${100 / cropHFrac}%`;
			img.style.left = `${-(l / PERMILLE) * (100 / cropWFrac)}%`;
			img.style.top = `${-(t / PERMILLE) * (100 / cropHFrac)}%`;
			img.style.maxWidth = "none";
			img.style.maxHeight = "none";
		} else {
			img.style.width = "100%";
			img.style.height = "100%";
			img.style.objectFit = "fill";
		}
	} else {
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "fill";
	}

	if (filter) img.style.filter = filter;
	if (opacity != null) img.style.opacity = String(opacity);

	wrap.appendChild(img);
	applyDuotoneFallback(wrap, pic);
	return wrap;
}

function buildFilterString(pic: PicShape): string {
	const parts: string[] = [];
	if (pic.grayscale) parts.push("grayscale(1)");
	if (pic.lumBrightPermille != null) {
		// <a:lum bright> is additive in PPTX; CSS brightness() is
		// multiplicative about 1. Map N per-mille → 1 + N/100000.
		parts.push(`brightness(${1 + pic.lumBrightPermille / PERMILLE})`);
	}
	if (pic.lumContrastPermille != null) {
		parts.push(`contrast(${1 + pic.lumContrastPermille / PERMILLE})`);
	}
	if (pic.biLevelPermille != null) {
		// Approximation: push contrast extremely high and nudge brightness
		// around the threshold so colors collapse toward black/white. A
		// faithful two-tone threshold needs an SVG <feComponentTransfer>
		// with a discrete table — TODO.
		const threshFrac = pic.biLevelPermille / PERMILLE;
		const bright = 1 + (0.5 - threshFrac);
		parts.push(`grayscale(1)`);
		parts.push(`brightness(${bright})`);
		parts.push(`contrast(1000)`);
	}
	return parts.join(" ");
}

// Best-effort duotone using mix-blend-mode on an overlay layer. A faithful
// implementation needs an inline SVG <filter> with feColorMatrix mapping
// luminance to a two-color ramp — TODO.
function applyDuotoneFallback(wrap: HTMLElement, pic: PicShape): void {
	if (!pic.duotone) return;
	wrap.style.backgroundColor = pic.duotone[0];
	const overlay = document.createElement("div");
	overlay.style.position = "absolute";
	overlay.style.inset = "0";
	overlay.style.backgroundColor = pic.duotone[1];
	overlay.style.mixBlendMode = "multiply";
	overlay.style.pointerEvents = "none";
	wrap.appendChild(overlay);
}

// Escape a URL for embedding inside a `url(...)` CSS function. Blob/data URLs
// from our open-xml-package loader don't contain unbalanced parens, but
// we still quote and escape defensively to avoid CSS injection vectors if
// the loader ever changes.
function cssUrl(src: string): string {
	// Encode backslashes, quotes, and newlines per CSS string rules.
	const escaped = src.replace(/(["\\])/g, "\\$1").replace(/\n/g, "\\A ");
	return `"${escaped}"`;
}
