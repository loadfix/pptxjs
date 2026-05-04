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
.${cls}-notes-slide {
	position: relative;
	/* width/height are set inline per-element from the deck's notesSz,
	   since notes pages can be any size the author chose (the PowerPoint
	   default is 7.5"x10" portrait, the inverse of a 4:3 slide). */
	margin: -16px auto 32px;
	background: #fdfdfd;
	border: 1px solid #d0d0d0;
	box-shadow: 0 1px 4px rgba(0,0,0,0.1);
	overflow: hidden;
	/* Visually subdue against the main slide so it reads as chrome. */
	opacity: 0.97;
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
