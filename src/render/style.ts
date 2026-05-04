export function makeStyleNode(cls: string, slideW: number, slideH: number): HTMLStyleElement {
	const style = document.createElement("style");
	style.textContent = `
.${cls}-slide {
	position: relative;
	width: ${slideW}px;
	height: ${slideH}px;
	margin: 0 auto 24px;
	background: #fff;
	box-shadow: 0 1px 4px rgba(0,0,0,0.2);
	overflow: hidden;
}
.${cls}-shape, .${cls}-pic, .${cls}-table {
	box-sizing: border-box;
}
`;
	return style;
}
