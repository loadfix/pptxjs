import { Presentation } from './presentation';
import { PresentationParser } from './presentation-parser';
import { HtmlRenderer } from './render';

export interface Options {
	className: string;
	inWrapper: boolean;
	debug: boolean;
	trimXmlDeclaration: boolean;
	useBase64URL: boolean;
	// Include slides that PowerPoint has marked hidden (<p:sldId show="0">).
	// Default false — hidden slides are dropped from the render list.
	showHidden: boolean;
	// When true, append a `.pptx-notes` div beneath each slide section
	// containing the speaker-notes text. Default false — notes are metadata
	// on the Presentation model and don't render unless explicitly asked for.
	renderNotes: boolean;
	// When true, render each slide's comments as absolutely-positioned pin
	// markers at the comment's (x,y) with hover tooltips showing the author
	// and body. Default false for the same reason as renderNotes.
	renderComments: boolean;
	// Optional callback invoked when a slide fails to parse. The slide still
	// takes its slot in `presentation.slides` with `parseError` set and will
	// render as a red banner, but hosts that want to surface/telemeter the
	// error (e.g. to their own logger) can hook in here. Errors thrown by
	// the callback itself are caught and warned — they never propagate back
	// into the load promise. Defaults to undefined (parse errors are logged
	// via console.warn only).
	onSlideError?: (slideIndex: number, err: unknown) => void;
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
	onSlideError: undefined,
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
	const pres = await parseAsync(data, userOptions);
	const nodes = await renderPresentation(pres, userOptions);

	styleContainer ??= bodyContainer;
	styleContainer.innerHTML = "";
	bodyContainer.innerHTML = "";

	for (const n of nodes) {
		const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
		c.appendChild(n);
	}

	return pres;
}
