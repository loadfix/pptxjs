import type { Slide } from '../presentation-parser';
import { emuToPx } from './geom';

// Append a `<div class="<cls>-notes">` after the slide section when the
// slide has speaker notes. Text content is set via textContent so arbitrary
// note text (including characters like "<") is safely escaped.
export function renderNotesBlock(slide: Slide, cls: string): HTMLElement | null {
	if (!slide.notes) return null;
	const div = document.createElement("div");
	div.className = `${cls}-notes`;
	div.dataset.slideIndex = String(slide.index);
	div.textContent = slide.notes.text;
	return div;
}

// Append pin markers for each comment, positioned absolutely at the
// comment's (x,y) in slide-local coordinates. Caller is responsible for
// placing the returned container inside a positioned parent (the slide
// section already is). Returns null when the slide has no comments.
export function renderCommentMarkers(slide: Slide, cls: string): HTMLElement | null {
	if (slide.comments.length === 0) return null;
	const layer = document.createElement("div");
	layer.className = `${cls}-comments`;
	// The layer itself is a point-events-free overlay filling the slide;
	// individual pins re-enable pointer events so their tooltips work.
	layer.style.position = "absolute";
	layer.style.left = "0";
	layer.style.top = "0";
	layer.style.width = "100%";
	layer.style.height = "100%";
	layer.style.pointerEvents = "none";

	for (const c of slide.comments) {
		const pin = document.createElement("div");
		pin.className = `${cls}-comment-pin`;
		pin.style.position = "absolute";
		pin.style.left = `${emuToPx(c.x)}px`;
		pin.style.top = `${emuToPx(c.y)}px`;
		pin.style.pointerEvents = "auto";
		pin.dataset.commentId = c.id;
		// Native title attribute gives us a free hover tooltip without
		// shipping a popover widget. "Author — text" matches PowerPoint's
		// comment bubble convention.
		const label = c.authorName ? `${c.authorName}: ${c.text}` : c.text;
		pin.title = label;
		pin.textContent = c.authorName ? c.authorName.charAt(0).toUpperCase() : "?";
		layer.appendChild(pin);
	}
	return layer;
}
