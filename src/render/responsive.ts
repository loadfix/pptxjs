/**
 * Responsive-mode glue: transforms `.${cls}-slide` sections to fit their
 * `.${cls}-slide-container` wrapper, keeping the slide's native authoring
 * pixel coordinates intact. Wiring lives here rather than in render/index.ts
 * so the observer can be loaded lazily / tree-shaken away when consumers
 * don't opt in.
 */

/**
 * Wrap a rendered slide `<section>` in a `.${cls}-slide-container` when the
 * caller asked for responsive layout. The container holds the slide's
 * aspect ratio so the vertical space is reserved correctly; the inner
 * `<section>` keeps its pixel size and is scaled by `attachResponsive` at
 * mount-time.
 */
export function wrapResponsive(
	section: HTMLElement,
	cls: string,
	slideW: number,
	slideH: number,
): HTMLElement {
	const wrapper = document.createElement("div");
	wrapper.className = `${cls}-slide-container`;
	wrapper.style.aspectRatio = `${slideW} / ${slideH}`;
	wrapper.appendChild(section);
	return wrapper;
}

/**
 * Attach a ResizeObserver to every `.${cls}-slide-container` under `root` so
 * each contained slide scales to the current container width. Called by
 * renderAsync after the rendered nodes are appended to the DOM — the
 * observer needs live layout sizes that only exist post-attach.
 *
 * Returns a disposer that disconnects the observer; callers that re-render
 * into the same container should call it to avoid dangling observers.
 */
export function attachResponsive(
	root: HTMLElement,
	slideWidthPx: number,
	className: string,
): () => void {
	const containers = root.querySelectorAll<HTMLElement>(`.${className}-slide-container`);
	if (containers.length === 0 || typeof ResizeObserver === "undefined") {
		return () => {};
	}
	const applyScale = (container: HTMLElement) => {
		const section = container.firstElementChild as HTMLElement | null;
		if (!section) return;
		const availableW = container.clientWidth;
		// Never scale up past 1.0 — slides are authored at their native
		// pixel size; upscaling blurs text. On a wide desktop the slide
		// simply centers inside its container.
		const scale = Math.min(1, availableW / slideWidthPx);
		section.style.transform = `scale(${scale})`;
	};
	const ro = new ResizeObserver((entries) => {
		for (const entry of entries) {
			applyScale(entry.target as HTMLElement);
		}
	});
	containers.forEach((c) => {
		applyScale(c);
		ro.observe(c);
	});
	return () => ro.disconnect();
}

/**
 * Responsive-mode CSS. Callers concatenate this onto the main stylesheet
 * only when `options.responsive === true`. The slide keeps its pixel-size
 * coordinate system (so authored shape positions still line up); the
 * container is the fluid outer box that clips the scaled slide.
 */
export function responsiveCss(cls: string, slideW: number): string {
	return `
.${cls}-slide-container {
	position: relative;
	width: 100%;
	max-width: ${slideW}px;
	margin: 0 auto 24px;
	overflow: hidden;
	background: #fff;
	box-shadow: 0 1px 4px rgba(0,0,0,0.2);
}
.${cls}-slide-container > .${cls}-slide {
	transform-origin: top left;
	margin: 0;
	box-shadow: none;
}
`;
}
