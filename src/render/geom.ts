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

// OOXML stores rotation in 60000ths of a degree (one full turn = 21 600 000).
// OOXML semantics are "flip first, then rotate around the flipped center".
// CSS applies transforms right-to-left, so listing rotate() BEFORE scale()
// means scale runs first (inner) and rotate runs last (outer) — matching OOXML.
export function transformStyle(
	rotation60k: number,
	flipH: boolean,
	flipV: boolean,
): { transform?: string; transformOrigin?: string } {
	const deg = (rotation60k || 0) / 60000;
	const parts: string[] = [];
	if (deg !== 0) parts.push(`rotate(${deg}deg)`);
	if (flipH) parts.push("scaleX(-1)");
	if (flipV) parts.push("scaleY(-1)");
	if (parts.length === 0) return {};
	return { transform: parts.join(" "), transformOrigin: "center center" };
}
