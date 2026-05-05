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
import { emuToPx, positionStyle } from './geom';
import { renderChart } from './chart';

export function renderChartFallback(shape: ChartFallbackShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	// Classes: the pptxjs-scoped "<cls>-chart" used for internal positioning
	// styles, plus the stable "pptx-chart" tag consumers can select on.
	// When the caller configured cls="pptx" the two collapse into a single
	// token, so we de-duplicate to keep class output tidy. data-kind mirrors
	// the class for selector-only consumers; data-chart-type carries the
	// probed series type when known.
	const scoped = `${cls}-chart`;
	wrap.className = scoped === "pptx-chart" ? "pptx-chart" : `${scoped} pptx-chart`;
	wrap.setAttribute("data-kind", "chart");
	if (shape.chartType) wrap.setAttribute("data-chart-type", shape.chartType);
	Object.assign(wrap.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));
	// Stable data-* hook for the conformance harness. W8-A3 will layer
	// data-chart-type on top once the chart XML is parsed.
	wrap.setAttribute("data-kind", "chart");
	if (shape.src) {
		const img = document.createElement("img");
		img.src = shape.src;
		if (shape.alt) img.setAttribute("alt", shape.alt);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "contain";
		wrap.appendChild(img);
	} else if (shape.model) {
		// Render from parsed ChartModel as inline SVG sized to the frame.
		const svg = renderChart(shape.model, emuToPx(shape.cx), emuToPx(shape.cy));
		svg.style.width = "100%";
		svg.style.height = "100%";
		wrap.appendChild(svg);
	} else {
		applyPlaceholderStyle(wrap, "[Chart]");
	}
	return wrap;
}

export function renderSmartArtFallback(shape: SmartArtFallbackShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	const scoped = `${cls}-smartart`;
	wrap.className = scoped === "pptx-smartart" ? "pptx-smartart" : `${scoped} pptx-smartart`;
	wrap.setAttribute("data-kind", "smartart");
	Object.assign(wrap.style, positionStyle(shape.x, shape.y, shape.cx, shape.cy));
	// Stable data-* hook for the conformance harness.
	wrap.setAttribute("data-kind", "smartart");
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
