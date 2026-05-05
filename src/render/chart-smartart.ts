// Fallback renderers for chart and SmartArt graphic frames.
//
// For charts, the parser has either resolved a cached preview image URL
// or left `src` null. We emit an <img> in the former case; in the latter
// we emit a bordered box with a centered "[Chart]" label so the slide
// doesn't look broken.
//
// For SmartArt, the parser normally replaces the frame with the expanded
// DrawingML shape list before reaching here — those shapes flow through
// the standard shape pipeline. This file handles the residual case where
// no drawing cache was found: we render a "[SmartArt]" placeholder.

import type { ChartFallbackShape, SmartArtFallbackShape } from '../graphic-frame';
import { positionStyle } from './geom';

export function renderChartFallback(shape: ChartFallbackShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-chart`;
	Object.assign(wrap.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));
	if (shape.src) {
		const img = document.createElement("img");
		img.src = shape.src;
		if (shape.alt) img.setAttribute("alt", shape.alt);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "contain";
		wrap.appendChild(img);
	} else {
		applyPlaceholderStyle(wrap, "[Chart]");
	}
	return wrap;
}

export function renderSmartArtFallback(shape: SmartArtFallbackShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-smartart`;
	Object.assign(wrap.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));
	applyPlaceholderStyle(wrap, "[SmartArt]");
	return wrap;
}

function applyPlaceholderStyle(el: HTMLElement, label: string): void {
	Object.assign(el.style, {
		border: "1px dashed #888",
		display: "flex",
		alignItems: "center",
		justifyContent: "center",
		color: "#666",
		font: "12px sans-serif",
		boxSizing: "border-box",
	});
	// textContent keeps the label out of innerHTML parsing; since it's a
	// fixed literal string this is belt-and-braces.
	el.textContent = label;
}
