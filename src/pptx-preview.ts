import { Presentation } from './presentation';
import { PresentationParser } from './presentation-parser';
import { HtmlRenderer } from './render';
import { attachResponsive } from './render/responsive';

// Re-export the structural metadata types so downstream consumers can type
// their own section/ToC UI without reaching into the parser module.
export type { Section, SlideSize } from './presentation-parser';

// Re-export the EMF converter override hook so UMD consumers (which can't
// rely on the bundler-driven dynamic import) can wire in a pre-loaded
// `emf-converter` module. The ESM build's code-splitting dynamic import
// path handles this automatically for bundled consumers; this escape hatch
// exists purely for script-tag / sandbox environments.
export { setEmfConverter } from './emf';

// EMU per CSS pixel at 96 DPI — duplicated here so pptx-preview.ts can
// compute the native slide width for attachResponsive without importing
// all of render.
const EMU_PER_PX = 9525;

export interface Options {
	className: string;
	inWrapper: boolean;
	debug: boolean;
	trimXmlDeclaration: boolean;
	useBase64URL: boolean;
	// Include slides that PowerPoint has marked hidden (<p:sldId show="0">).
	// Default false — hidden slides are dropped from the render list.
	showHidden: boolean;
	// When true, append a `.pptx-notes-slide` section beneath each slide
	// section containing the notesSlide rendered through the same shape
	// pipeline (bullets, colours, placeholders, images all honoured).
	// Default false — notes are metadata on the Presentation model and
	// don't render unless explicitly asked for.
	renderNotes: boolean;
	// When true, render each slide's comments as absolutely-positioned pin
	// markers at the comment's (x,y) with hover tooltips showing the author
	// and body. Default false for the same reason as renderNotes.
	renderComments: boolean;
	// When true, append one or more `.pptx-handout-page` sections after the
	// slide list, each showing the deck's slides as scaled-down thumbnails
	// overlaid on the handoutMaster's chrome (header/footer/date/slideNum
	// placeholders). Six slides per page in a 3×2 grid. Default false —
	// handouts are a print-layout feature most viewers don't want inline.
	renderHandouts: boolean;
	// Optional callback invoked when a slide fails to parse. The slide still
	// takes its slot in `presentation.slides` with `parseError` set and will
	// render as a red banner, but hosts that want to surface/telemeter the
	// error (e.g. to their own logger) can hook in here. Errors thrown by
	// the callback itself are caught and warned — they never propagate back
	// into the load promise. Defaults to undefined (parse errors are logged
	// via console.warn only).
	onSlideError?: (slideIndex: number, err: unknown) => void;
	// When true (default), EMF / WMF images embedded in the deck are
	// rasterised to PNG via `emf-converter` (lazy-loaded) and surfaced as
	// normal `<img>` sources so brand logos and Illustrator-exported
	// diagrams render. When false, EMF/WMF images are skipped — the image
	// slot renders empty. Useful in environments that reject the extra
	// bundle chunk or canvas dependency deterministically.
	convertEmf: boolean;
	/**
	 * Responsive / mobile rendering. When `true`, each slide is wrapped in
	 * a `<div class="pptx-slide-container">` that applies a CSS
	 * `transform: scale(...)` sized to fit the container's width, using a
	 * ResizeObserver to update the scale as the viewport changes. Slide
	 * aspect ratio is preserved; content inside the slide keeps its
	 * pixel coordinates so shape positioning stays exact.
	 *
	 * Default `false` — existing consumers keep the fixed-size slide
	 * layout they rendered before this option landed.
	 */
	responsive?: boolean;
}

export const defaultOptions: Options = {
	className: "pptx",
	inWrapper: true,
	debug: false,
	trimXmlDeclaration: true,
	useBase64URL: false,
	showHidden: false,
	renderNotes: false,
	renderComments: false,
	renderHandouts: false,
	onSlideError: undefined,
	convertEmf: true,
	responsive: false,
};

function mergeOptions(userOptions?: Partial<Options>): Options {
	return { ...defaultOptions, ...userOptions };
}

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<Presentation> {
	const ops = mergeOptions(userOptions);
	return Presentation.load(data, new PresentationParser(ops), ops);
}

export async function renderPresentation(presentation: Presentation, userOptions?: Partial<Options>): Promise<Node[]> {
	const ops = mergeOptions(userOptions);
	const renderer = new HtmlRenderer();
	return await renderer.render(presentation, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<Presentation> {
	const ops = mergeOptions(userOptions);
	const pres = await parseAsync(data, ops);
	const nodes = await renderPresentation(pres, ops);

	styleContainer ??= bodyContainer;
	styleContainer.innerHTML = "";
	bodyContainer.innerHTML = "";

	for (const n of nodes) {
		const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
		c.appendChild(n);
	}

	// Responsive mode needs a live ResizeObserver so each `.pptx-slide-container`
	// recomputes its scale factor on container resize. Must run after the
	// nodes are connected — the observer needs real layout sizes.
	if (ops.responsive) {
		const slideWPx = pres.slideSize.cx / EMU_PER_PX;
		attachResponsive(bodyContainer, slideWPx, ops.className);
	}

	return pres;
}
