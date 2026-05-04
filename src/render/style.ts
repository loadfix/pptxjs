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
.${cls}-notes {
	width: ${slideW}px;
	margin: -16px auto 24px;
	padding: 12px 16px;
	background: #fafafa;
	border: 1px solid #e0e0e0;
	border-top: none;
	font: 13px/1.5 system-ui, sans-serif;
	color: #333;
	white-space: pre-wrap;
}
.${cls}-comment-pin {
	width: 20px;
	height: 20px;
	border-radius: 50%;
	background: #ffd54f;
	border: 1px solid #b28704;
	color: #3d2b00;
	font: bold 11px/18px system-ui, sans-serif;
	text-align: center;
	cursor: help;
	box-shadow: 0 1px 2px rgba(0,0,0,0.3);
	transform: translate(-50%, -50%);
}
`;
	return style;
}
