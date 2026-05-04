import type { PicShape } from '../presentation-parser';
import { positionStyle } from './geom';

export function renderPic(pic: PicShape, cls: string): HTMLElement {
	const wrap = document.createElement("div");
	wrap.className = `${cls}-pic`;
	Object.assign(wrap.style, positionStyle(pic.x, pic.y, pic.cx, pic.cy));
	if (pic.src) {
		const img = document.createElement("img");
		img.src = pic.src;
		// The DOCX-side render guidance warns against putting document-derived
		// strings into innerHTML/CSS; `alt`/`title` via setAttribute is safe
		// (the browser HTML-encodes attribute values).
		// PicShape.alt already falls back descr → name → "" during parse, so
		// it's the right value for <img alt=>. Only emit `title` when the
		// source set an explicit `title` attribute on the picture.
		if (pic.alt) img.setAttribute("alt", pic.alt);
		if (pic.title) img.setAttribute("title", pic.title);
		img.style.width = "100%";
		img.style.height = "100%";
		img.style.objectFit = "fill";
		wrap.appendChild(img);
	}
	return wrap;
}
