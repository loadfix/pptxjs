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
/* Handout-master "print layout" page. Sized inline per-element from the
   deck's notesSz (OOXML uses one <p:notesSz> for both notes *and*
   handouts). The master's chrome shapes are appended as direct children
   with absolute positioning from their parsed EMU coords; the thumbnail
   grid overlays it as a positioned block.

   The precise position of the six thumbnails on a handout page is a
   PowerPoint print-time concern, not an OOXML-specified one — the
   handoutMaster only carries header/footer/date/slideNum placeholders.
   We synthesize a 3-row × 2-col CSS grid here with generous padding so
   the printed template chrome still shows around it. */
.${cls}-handout-page {
	position: relative;
	margin: 0 auto 32px;
	background: #fff;
	box-shadow: 0 1px 4px rgba(0,0,0,0.2);
	overflow: hidden;
}
.${cls}-handout-grid {
	position: absolute;
	/* Leave ~10% margin on each edge for the master's chrome. The values
	   are deliberately coarse — the chrome placeholders (header/footer)
	   typically sit within the top and bottom inches of a Letter page. */
	top: 10%;
	left: 8%;
	right: 8%;
	bottom: 10%;
	display: grid;
	grid-template-columns: repeat(2, 1fr);
	grid-template-rows: repeat(3, 1fr);
	gap: 12px;
}
.${cls}-handout-cell {
	position: relative;
	border: 1px solid #d0d0d0;
	background: #fff;
	overflow: hidden;
}
.${cls}-handout-thumb {
	position: absolute;
	top: 0;
	left: 0;
	/* Each cell is sized by grid; the thumbnail inside is the full-size
	   slide markup, scaled down via transform so shapes stay crisp.
	   The scale factor matches the slide's natural width against the
	   cell's width at typical handout sizes: ~24% for a 9144000×6858000
	   slide tiled into the grid above. Hosts that need an exact fit can
	   override this rule; precise math isn't load-bearing here since
	   the surrounding cell has overflow:hidden and the grid is a
	   synthesised layout anyway. */
	transform: scale(0.24);
	transform-origin: top left;
}
.${cls}-handout-thumb .${cls}-slide {
	box-shadow: none;
	margin: 0;
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
