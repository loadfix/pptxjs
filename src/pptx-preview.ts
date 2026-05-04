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
}

export const defaultOptions: Options = {
	className: "pptx",
	inWrapper: true,
	debug: false,
	trimXmlDeclaration: true,
	useBase64URL: false,
	showHidden: false,
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
