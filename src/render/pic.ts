import type { PicShape } from '../presentation-parser';
import { positionStyle, transformStyle } from './geom';

// Per-mille denominator used throughout OOXML blip-effect attributes
// (srcRect/alphaModFix/lum/biLevel/tile). 100000 = 100%.
const PERMILLE = 100000;

export function renderPic(pic: PicShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	Object.assign(wrap.style, transformStyle(pic.rotation60k, pic.flipH, pic.flipV));

	if (!pic.src) return wrap;

	// Build the CSS filter list once — applies to both stretch and tile
	// rendering paths.
	const filter = buildFilterString(pic);
	const opacity = pic.alphaPermille != null ? pic.alphaPermille / PERMILLE : null;

	if (!pic.stretch) {
		// Tile fill — render as a repeating background on the container.
		wrap.style.backgroundImage = `url(${cssUrl(pic.src)})`;
		wrap.style.backgroundRepeat = "repeat";
		wrap.style.backgroundSize = "auto";
		if (pic.tile) {
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
	// The DOCX-side render guidance warns against putting document-derived
	// strings into innerHTML/CSS; `alt` via setAttribute is safe (the
	// browser HTML-encodes attribute values).
	if (pic.alt) img.setAttribute("alt", pic.alt);

	if (pic.srcRectPermille) {
		// Crop via overflow-hidden wrapper + scaled/offset <img>.
		const { l, t, r, b } = pic.srcRectPermille;
		const cropWFrac = 1 - (l + r) / PERMILLE;
		const cropHFrac = 1 - (t + b) / PERMILLE;
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
		parts.push(`brightness(${1 + pic.lumBrightPermille / PERMILLE})`);
	}
	if (pic.lumContrastPermille != null) {
		parts.push(`contrast(${1 + pic.lumContrastPermille / PERMILLE})`);
	}
	if (pic.biLevelPermille != null) {
		const threshFrac = pic.biLevelPermille / PERMILLE;
		const bright = 1 + (0.5 - threshFrac);
		parts.push(`grayscale(1)`);
		parts.push(`brightness(${bright})`);
		parts.push(`contrast(1000)`);
	}
	return parts.join(" ");
}

// Best-effort duotone via a mix-blend-mode overlay. A faithful implementation
// needs an inline SVG <filter> with feColorMatrix mapping luminance to a
// two-color ramp — TODO.
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

function cssUrl(src: string): string {
	const escaped = src.replace(/(["\\])/g, "\\$1").replace(/\n/g, "\\A ");
	return `"${escaped}"`;
}
