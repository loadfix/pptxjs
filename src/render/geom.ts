// English Metric Units per pixel at 96 DPI. 914400 EMU = 1 inch = 96 px.
export const EMU_PER_PX = 9525;
export const emuToPx = (emu: number) => emu / EMU_PER_PX;

export const SVG_NS = "http://www.w3.org/2000/svg";

export function positionStyle(x: number, y: number, cx: number, cy: number): Partial<CSSStyleDeclaration> {
	return {
		position: "absolute",
		left: `${emuToPx(x)}px`,
		top: `${emuToPx(y)}px`,
		width: `${emuToPx(cx)}px`,
		height: `${emuToPx(cy)}px`,
	};
}
