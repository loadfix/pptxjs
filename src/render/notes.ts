import type { Slide } from '../presentation-parser';
import type { Presentation } from '../presentation';
import { emuToPx } from './geom';
import { renderShapeLike } from './dispatch';

// Append a `<section class="<cls>-notes-slide">` after the slide section when
// the slide has speaker notes. The section holds the notes-slide's shapes
// rendered through the main dispatch pipeline, so paragraph structure,
// bullets, inherited colors from the notesMaster, and any embedded images
// render with full fidelity (rather than as a flat text blob).
//
// Layout: the notes section is sized to the presentation's <p:notesSz> and
// positioned below the main slide by the shared stylesheet (margin-top:
// -16px matches the old .pptx-notes overlap so the notes page appears to
// dock under its slide).
export function renderNotesBlock(
	slide: Slide,
	cls: string,
	presentation: Presentation,
	embedUrls: Map<string, string>,
): HTMLElement | null {
	if (!slide.notes) return null;
	const section = document.createElement("section");
	// `pptx-notes` is a stable hook the ooxml-validate corpus asserts on
	// (slide-notes case); `${cls}-notes-slide` retains the deck-scoped class
	// so the shared stylesheet's chrome rules still match.
	section.className = `${cls}-notes-slide pptx-notes`;
	section.dataset.notes = "true";
	section.dataset.slideIndex = String(slide.index);
	// Inline size — the style block declares the box chrome (border, margins)
	// but each deck's notesSz may differ, so width/height are set per-element.
	// PowerPoint's default 6858000 × 9144000 EMU is 7.5"×10".
	section.style.width = `${emuToPx(presentation.notesSize.cx)}px`;
	section.style.height = `${emuToPx(presentation.notesSize.cy)}px`;

	// Field substitution on notes shapes resolves slidenum / header / footer /
	// datetime against the owning slide (not the notes slide itself), since
	// that's what a speaker-notes page traditionally reads.
	const fieldCtx = { slide, firstSlideNum: presentation.firstSlideNum };
	for (const shape of slide.notes.shapes) {
		// `hyperlinkUrls` is the owning slide's map — notes slides rarely
		// carry their own links, and any rIds on notes shapes won't collide
		// with slide-level ones because parse happens against the notes part.
		// We pass the slide's map as a best-effort fallback; unresolved rIds
		// simply render without a link, same as on a regular slide.
		const node = renderShapeLike(
			shape,
			cls,
			embedUrls,
			presentation.tableStyles,
			slide.hyperlinkUrls,
			fieldCtx,
		);
		if (node) section.appendChild(node);
	}
	return section;
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
