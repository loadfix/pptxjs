import type { Slide } from '../presentation-parser';

export function applyBackground(section: HTMLElement, slide: Slide): void {
	if (slide.background?.kind === 'solid') {
		section.style.background = slide.background.colorHex;
	}
}
