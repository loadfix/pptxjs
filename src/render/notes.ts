import type { Slide } from '../presentation-parser';
import type { NotesMasterBundle } from '../presentation';
import { resolveTypeface, resolveSchemeClr } from '../theme';
import { emuToPx } from './geom';

// Append a `<div class="<cls>-notes">` after the slide section when the
// slide has speaker notes. Text content is set via textContent so arbitrary
// note text (including characters like "<") is safely escaped.
//
// When a notesMaster bundle is available, the level-0 run defaults from its
// <p:notesStyle> (font family, size, color) are applied inline to the div so
// speaker notes pick up the deck's intended typography instead of browser
// defaults. This is a minimum-viable typography pass — we don't do per-run
// styling because P12 only captures plain text from the notesSlide; richer
// styling belongs to a later upgrade that traverses the notesSlide's shapes.
export function renderNotesBlock(
	slide: Slide,
	cls: string,
	notesMaster: NotesMasterBundle | null = null,
): HTMLElement | null {
	if (!slide.notes) return null;
	const div = document.createElement("div");
	div.className = `${cls}-notes`;
	div.dataset.slideIndex = String(slide.index);
	div.textContent = slide.notes.text;

	if (notesMaster) {
		applyNotesMasterTypography(div, notesMaster);
	}
	return div;
}

// Pull the notes master's first-level body defaults and convert them into
// CSS on the notes div. We read `notesStyle[0]` (the lvl1 defaults) because
// the concatenated notes text in NotesSlide collapses all paragraphs into a
// single string — we have no level information to branch on yet. Any field
// left null on the defRPr is left alone so the .pptx-notes base stylesheet
// still governs it.
function applyNotesMasterTypography(div: HTMLElement, notesMaster: NotesMasterBundle): void {
	const lvl0 = notesMaster.notesStyle[0];
	const run = lvl0?.run;
	if (!run) return;

	if (run.fontFamily) {
		// Theme-ref typefaces (+mj-lt / +mn-lt) resolve against the notes
		// master's own theme. A null resolution means the typeface wasn't in
		// the theme's fontScheme — fall through to the literal value.
		const resolved = resolveTypeface(run.fontFamily, notesMaster.theme) ?? run.fontFamily;
		div.style.fontFamily = resolved;
	}
	if (run.sizeHundredths != null) {
		// sz is in hundredths of a point. Convert to pt so CSS understands it;
		// the browser handles pt→px scaling.
		div.style.fontSize = `${run.sizeHundredths / 100}pt`;
	}
	// Resolve color. The notesStyle defRPr typically holds a schemeClr
	// (e.g. tx1) rather than an srgbClr, so we need to run it through the
	// notes master's own clrMap + theme. parseRunProps already attempted
	// this once but only when the level-styles parser was given a clrMap
	// (it currently isn't), so we may still see a schemeSlot with no hex.
	const colorHex = run.colorHex
		?? (run.colorSchemeSlot
			? resolveSchemeClr(run.colorSchemeSlot, notesMaster.clrMap, notesMaster.theme)
			: null);
	if (colorHex) div.style.color = colorHex;
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
