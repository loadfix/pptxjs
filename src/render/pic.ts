import type { PicShape } from '../presentation-parser';
import { positionStyle } from './geom';
import { wrapInHyperlink } from './hyperlink';

export function renderPic(pic: PicShape, cls: string, hyperlinkUrls: Map<string, string>): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	if (pic.src) {
		const img = document.createElement("img");
		img.src = pic.src;
		// The DOCX-side render guidance warns against putting document-derived
		// strings into innerHTML/CSS; `alt` via setAttribute is safe (the
		// browser HTML-encodes attribute values).
		if (pic.alt) img.setAttribute("alt", pic.alt);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "fill";
		wrap.appendChild(img);
	}
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
