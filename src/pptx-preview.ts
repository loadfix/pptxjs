import { Presentation } from './presentation';
import { PresentationParser } from './presentation-parser';
import { HtmlRenderer, attachResponsive } from './html-renderer';

// EMU per CSS pixel at 96 DPI — duplicated here so pptx-preview.ts can
// compute the native slide width for attachResponsive without importing
// all of html-renderer.
const EMU_PER_PX = 9525;

export interface Options {
	className: string;
	inWrapper: boolean;
	debug: boolean;
	trimXmlDeclaration: boolean;
	useBase64URL: boolean;
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
