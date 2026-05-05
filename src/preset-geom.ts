// OOXML preset shape geometry → SVG path conversion.
//
// PowerPoint stores most shapes as a preset name (e.g. "roundRect",
// "rightArrow") plus an optional adjust-value list controlling the silhouette.
// The full set is defined in ECMA-376 §20.1.9 as `prstGeomDefs`; there are
// ~180 of them and implementing each with full formula fidelity is a lot of
// code. We cover the ~40 most common presets with reasonable approximations
// so that typical real-world decks render recognisably.
//
// All paths are emitted in a local coordinate space of (0..w, 0..h) where
// w/h are the shape's EMU extents, so the caller can drop them into a
// `viewBox="0 0 w h" preserveAspectRatio="none"` SVG — exactly how custGeom
// paths are rendered.
//
// Adjust values (<a:gd name="adj1" fmla="val N"/>) are in per-100000 units.
// For most presets only adj1..adj3 matter; we ignore higher ones.

type Adj = Map<string, number> | undefined;

function getAdj(av: Adj, name: string, fallback: number): number {
	if (!av) return fallback;
	const v = av.get(name);
	return v == null ? fallback : v;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const DEBUG = typeof (globalThis as any).process !== 'undefined' && !!((globalThis as any).process as { env?: Record<string, string> }).env?.DEBUG_PPTXJS;

/**
 * Convert an OOXML preset-geometry name + adjust-value list to an SVG `d`
 * attribute expressed in (0..w, 0..h) local coordinates. Returns null for
 * unknown presets; caller should fall back to a plain rectangle.
 */
export function presetToSvgPath(preset: string, w: number, h: number, avLst?: Map<string, number>): string | null {
	const w0 = Math.max(w, 1);
	const h0 = Math.max(h, 1);
	const min = Math.min(w0, h0);

	switch (preset) {
		case "rect":
		case "flowChartProcess":
			return rectPath(w0, h0);

		case "roundRect":
		case "flowChartTerminator": {
			// adj1 ∈ [0, 50000] — corner radius as per-100000 of shorter side.
			const adj1 = getAdj(avLst, "adj1", preset === "flowChartTerminator" ? 50000 : 16667);
			const r = Math.min((adj1 / 100000) * min, min / 2);
			return roundRectPath(w0, h0, r);
		}

		case "ellipse":
		case "flowChartConnector":
			return ellipsePath(w0, h0);

		case "triangle": {
			// adj1 = x of apex, as per-100000 of w. Default 50000 (isoceles).
			const adj1 = getAdj(avLst, "adj1", 50000);
			const apexX = (adj1 / 100000) * w0;
			return `M ${apexX} 0 L ${w0} ${h0} L 0 ${h0} Z`;
		}

		case "rtTriangle":
			return `M 0 0 L 0 ${h0} L ${w0} ${h0} Z`;

		case "diamond":
		case "flowChartDecision":
			return `M ${w0 / 2} 0 L ${w0} ${h0 / 2} L ${w0 / 2} ${h0} L 0 ${h0 / 2} Z`;

		case "parallelogram":
		case "flowChartData": {
			// adj1 = slant as per-100000 of w; ≤100000 for normal parallelograms.
			const adj1 = getAdj(avLst, "adj1", 25000);
			const s = Math.min(Math.max(adj1 / 100000, 0), 1) * w0;
			return `M ${s} 0 L ${w0} 0 L ${w0 - s} ${h0} L 0 ${h0} Z`;
		}

		case "trapezoid": {
			// adj1 = inset of top edges, per-100000 of w; 0..50000.
			const adj1 = getAdj(avLst, "adj1", 25000);
			const s = Math.min(Math.max(adj1 / 100000, 0), 0.5) * w0;
			return `M ${s} 0 L ${w0 - s} 0 L ${w0} ${h0} L 0 ${h0} Z`;
		}

		case "pentagon":
			return regularPolygonPath(w0, h0, 5, -Math.PI / 2);

		case "hexagon":
			return regularPolygonPath(w0, h0, 6, 0);

		case "octagon":
			return regularPolygonPath(w0, h0, 8, Math.PI / 8);

		case "rightArrow":
			return arrowPath(w0, h0, avLst, "right");
		case "leftArrow":
			return arrowPath(w0, h0, avLst, "left");
		case "upArrow":
			return arrowPath(w0, h0, avLst, "up");
		case "downArrow":
			return arrowPath(w0, h0, avLst, "down");
		case "leftRightArrow":
			return leftRightArrowPath(w0, h0, avLst);

		case "bentArrow":
			return bentArrowPath(w0, h0, avLst);
		case "curvedRightArrow":
			return curvedRightArrowPath(w0, h0, avLst);

		case "star5":
			return starPath(w0, h0, 5, avLst, 38000);
		case "star6":
			return starPath(w0, h0, 6, avLst, 28000);
		case "star7":
			return starPath(w0, h0, 7, avLst, 34500);
		case "star8":
			return starPath(w0, h0, 8, avLst, 37500);
		case "star10":
			return starPath(w0, h0, 10, avLst, 42500);
		case "star12":
			return starPath(w0, h0, 12, avLst, 37500);
		case "star16":
			return starPath(w0, h0, 16, avLst, 37500);

		case "chevron": {
			// adj1 = notch depth, per-100000 of w; default 50000.
			const adj1 = getAdj(avLst, "adj1", 50000);
			const d = (adj1 / 100000) * w0;
			return `M 0 0 L ${w0 - d} 0 L ${w0} ${h0 / 2} L ${w0 - d} ${h0} L 0 ${h0} L ${d} ${h0 / 2} Z`;
		}

		case "homePlate": {
			const adj1 = getAdj(avLst, "adj1", 50000);
			const d = (adj1 / 100000) * w0;
			return `M 0 0 L ${w0 - d} 0 L ${w0} ${h0 / 2} L ${w0 - d} ${h0} L 0 ${h0} Z`;
		}

		case "plaque": {
			// adj1 = corner cutoff depth, per-100000 of shorter side.
			const adj1 = getAdj(avLst, "adj1", 16667);
			const c = Math.min((adj1 / 100000) * min, min / 2);
			return `M 0 ${c} Q 0 0 ${c} 0 L ${w0 - c} 0 Q ${w0} 0 ${w0} ${c} L ${w0} ${h0 - c} Q ${w0} ${h0} ${w0 - c} ${h0} L ${c} ${h0} Q 0 ${h0} 0 ${h0 - c} Z`;
		}

		case "noSmoking": {
			// Circle with a diagonal bar. We approximate with two concentric
			// circles (ring) plus a thick rotated rectangle through the centre
			// as a single even-odd fill — but SVG `d` can't set fill-rule
			// inline, so emit the outer circle + the bar as two subpaths.
			const cx = w0 / 2, cy = h0 / 2;
			const rx = w0 / 2, ry = h0 / 2;
			const barThickness = getAdj(avLst, "adj1", 18750) / 100000 * min;
			// Outer circle (counter-clockwise for even-odd cutout), inner circle (clockwise)
			const inner = Math.max(Math.min(rx, ry) - barThickness * 0.6, 1);
			const ring = ellipseDWithHole(cx, cy, rx, ry, inner, inner);
			// Diagonal bar (rotated 45deg rectangle centred on (cx,cy))
			const half = barThickness / 2;
			const r45 = Math.SQRT1_2;
			const dx = half * r45, dy = half * r45;
			const len = Math.min(rx, ry) * 2;
			const hx = (len / 2) * r45, hy = (len / 2) * r45;
			const bar =
				`M ${cx - hx - dx} ${cy - hy + dy} ` +
				`L ${cx + hx - dx} ${cy + hy + dy} ` +
				`L ${cx + hx + dx} ${cy + hy - dy} ` +
				`L ${cx - hx + dx} ${cy - hy - dy} Z`;
			return `${ring} ${bar}`;
		}

		case "plus": {
			// adj1 = bar thickness, per-100000 of shorter side. Default 25000.
			const adj1 = getAdj(avLst, "adj1", 25000);
			const t = Math.min(Math.max(adj1 / 100000, 0), 0.5);
			const hx = t * w0, hy = t * h0;
			return `M ${hx} 0 L ${w0 - hx} 0 L ${w0 - hx} ${hy} L ${w0} ${hy} L ${w0} ${h0 - hy} L ${w0 - hx} ${h0 - hy} L ${w0 - hx} ${h0} L ${hx} ${h0} L ${hx} ${h0 - hy} L 0 ${h0 - hy} L 0 ${hy} L ${hx} ${hy} Z`;
		}

		case "line":
		case "straightConnector1":
			return `M 0 0 L ${w0} ${h0}`;

		case "callout1":
		case "callout2":
		case "wedgeRectCallout":
			return wedgeRectCallout(w0, h0, avLst, 0);
		case "wedgeRoundRectCallout":
			return wedgeRectCallout(w0, h0, avLst, 16667);
		case "wedgeEllipseCallout":
			return wedgeEllipseCallout(w0, h0, avLst);

		default:
			if (DEBUG) console.warn(`[pptxjs] unsupported prstGeom "${preset}" — falling back to rectangle`);
			return null;
	}
}

function rectPath(w: number, h: number): string {
	return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
}

function roundRectPath(w: number, h: number, r: number): string {
	if (r <= 0) return rectPath(w, h);
	return `M ${r} 0 L ${w - r} 0 Q ${w} 0 ${w} ${r} L ${w} ${h - r} Q ${w} ${h} ${w - r} ${h} L ${r} ${h} Q 0 ${h} 0 ${h - r} L 0 ${r} Q 0 0 ${r} 0 Z`;
}

function ellipsePath(w: number, h: number): string {
	const rx = w / 2, ry = h / 2;
	// Full ellipse via two 180° arcs.
	return `M 0 ${ry} A ${rx} ${ry} 0 1 0 ${w} ${ry} A ${rx} ${ry} 0 1 0 0 ${ry} Z`;
}

function ellipseDWithHole(cx: number, cy: number, rx: number, ry: number, rx2: number, ry2: number): string {
	// Outer ellipse clockwise, inner ellipse counter-clockwise (nonzero rule
	// produces a ring when sub-paths have opposite orientations).
	const outer = `M ${cx - rx} ${cy} A ${rx} ${ry} 0 1 0 ${cx + rx} ${cy} A ${rx} ${ry} 0 1 0 ${cx - rx} ${cy} Z`;
	const inner = `M ${cx - rx2} ${cy} A ${rx2} ${ry2} 0 1 1 ${cx + rx2} ${cy} A ${rx2} ${ry2} 0 1 1 ${cx - rx2} ${cy} Z`;
	return `${outer} ${inner}`;
}

// Regular n-gon inscribed in the box, optionally rotated. For pentagon/hexagon
// we inscribe in a circle of radius min(w,h)/2 then scale X,Y to fit the box.
function regularPolygonPath(w: number, h: number, n: number, rotate: number): string {
	const cx = w / 2, cy = h / 2;
	const rx = w / 2, ry = h / 2;
	const pts: string[] = [];
	for (let i = 0; i < n; i++) {
		const ang = rotate + (i * 2 * Math.PI) / n;
		const x = cx + rx * Math.cos(ang);
		const y = cy + ry * Math.sin(ang);
		pts.push(`${x} ${y}`);
	}
	return `M ${pts[0]} ` + pts.slice(1).map(p => `L ${p}`).join(" ") + " Z";
}

// Generic arrow. `dir` picks orientation; we build a right-facing arrow in
// a unit box then rotate/flip point coords.
//
// adj1 = shaft thickness as per-100000 of shorter side (height of shaft)
// adj2 = head length as per-100000 of longer side
// Defaults chosen so the shape looks like PowerPoint's default.
function arrowPath(w: number, h: number, av: Adj, dir: "right" | "left" | "up" | "down"): string {
	const adj1 = getAdj(av, "adj1", 50000); // shaft thickness fraction
	const adj2 = getAdj(av, "adj2", 50000); // head length fraction
	// Build right-facing arrow in w×h box.
	const shaft = Math.min(Math.max(adj1 / 100000, 0), 1);
	const headLen = Math.min(Math.max(adj2 / 100000, 0), 1);
	const isH = dir === "right" || dir === "left";
	const main = isH ? w : h;
	const cross = isH ? h : w;
	const st = ((1 - shaft) / 2) * cross;   // shaft top (offset from edge)
	const sb = cross - st;                  // shaft bottom
	const hl = headLen * main;              // head length along main axis
	// Points in a right-facing "local" arrow: (along, cross).
	const pts: [number, number][] = [
		[0, st],
		[main - hl, st],
		[main - hl, 0],
		[main, cross / 2],
		[main - hl, cross],
		[main - hl, sb],
		[0, sb],
	];
	return mapArrowPoints(pts, dir, w, h);
}

// Convert arrow "along/cross" points to (x,y) in w,h, honouring direction.
function mapArrowPoints(pts: [number, number][], dir: "right" | "left" | "up" | "down", w: number, h: number): string {
	const mapped = pts.map(([a, c]): [number, number] => {
		switch (dir) {
			case "right": return [a, c];
			case "left": return [w - a, c];
			case "down": return [c, a];
			case "up": return [c, h - a];
		}
	});
	return `M ${mapped[0][0]} ${mapped[0][1]} ` +
		mapped.slice(1).map(([x, y]) => `L ${x} ${y}`).join(" ") + " Z";
}

function leftRightArrowPath(w: number, h: number, av: Adj): string {
	const adj1 = getAdj(av, "adj1", 50000); // shaft thickness
	const adj2 = getAdj(av, "adj2", 25000); // each head length fraction of w
	const shaft = Math.min(Math.max(adj1 / 100000, 0), 1);
	const head = Math.min(Math.max(adj2 / 100000, 0), 0.5);
	const st = ((1 - shaft) / 2) * h;
	const sb = h - st;
	const hl = head * w;
	const pts: [number, number][] = [
		[0, h / 2],
		[hl, 0],
		[hl, st],
		[w - hl, st],
		[w - hl, 0],
		[w, h / 2],
		[w - hl, h],
		[w - hl, sb],
		[hl, sb],
		[hl, h],
	];
	return `M ${pts[0][0]} ${pts[0][1]} ` +
		pts.slice(1).map(([x, y]) => `L ${x} ${y}`).join(" ") + " Z";
}

// Bent (L-shaped) arrow. Vertical shaft on the left rising from the bottom,
// bends to a horizontal shaft with the arrow head at the right edge.
// Approximated with straight segments.
function bentArrowPath(w: number, h: number, av: Adj): string {
	const shaftFrac = getAdj(av, "adj1", 25000) / 100000;
	const headFrac = getAdj(av, "adj2", 25000) / 100000;
	const mn = Math.min(w, h);
	const t = shaftFrac * mn;                 // shaft thickness
	const hl = Math.min(headFrac * w, w - t); // head length along x
	const hw = Math.min(t * 2.2, h);          // head cross width
	const yBend = h / 2;
	const yShaftTop = yBend - t / 2;
	const yShaftBot = yBend + t / 2;
	const xHeadBase = Math.max(w - hl, t);
	const yHeadTop = yBend - hw / 2;
	const yHeadBot = yBend + hw / 2;
	return `M 0 ${h} L 0 ${yShaftTop} L ${xHeadBase} ${yShaftTop} L ${xHeadBase} ${yHeadTop} L ${w} ${yBend} L ${xHeadBase} ${yHeadBot} L ${xHeadBase} ${yShaftBot} L ${t} ${yShaftBot} L ${t} ${h} Z`;
}

// Curved right arrow — approximate as a quarter-annulus from (0,h) sweeping
// clockwise up to (w,0) with an arrowhead at the end. Exact OOXML curves are
// complex so we use two arcs.
function curvedRightArrowPath(w: number, h: number, av: Adj): string {
	const shaftFrac = getAdj(av, "adj1", 25000) / 100000;
	const headFrac = getAdj(av, "adj3", 25000) / 100000;
	const t = Math.min(w, h) * shaftFrac; // shaft thickness
	const hl = w * headFrac;
	// Outer arc: from (0, h) along ellipse (rx=w, ry=h) to (w-hl, 0).
	// Inner arc: from (w-hl, t) back to (0, h-t) along ellipse (rx=w-t, ry=h-t).
	const rxO = w, ryO = h;
	const rxI = Math.max(w - t, 1), ryI = Math.max(h - t, 1);
	const tipX = w, tipY = 0;
	const headBaseX = w - hl;
	return `M 0 ${h} A ${rxO} ${ryO} 0 0 1 ${headBaseX} 0 L ${headBaseX} ${-hl / 2} L ${tipX} ${tipY + hl / 2} L ${headBaseX} ${hl} L ${headBaseX} ${t} A ${rxI} ${ryI} 0 0 0 ${t} ${h - t} L ${t} ${h} Z`;
}

// N-pointed star. adj1 = inner-radius as per-100000 of outer (≈ 0.38 for star5).
function starPath(w: number, h: number, points: number, av: Adj, defaultInner: number): string {
	const inner = getAdj(av, "adj1", defaultInner) / 100000;
	const cx = w / 2, cy = h / 2;
	const rxO = w / 2, ryO = h / 2;
	const rxI = rxO * inner, ryI = ryO * inner;
	const pts: string[] = [];
	const start = -Math.PI / 2; // point up
	const step = Math.PI / points; // angle between successive outer+inner verts
	for (let i = 0; i < points * 2; i++) {
		const ang = start + i * step;
		const isOuter = i % 2 === 0;
		const rx = isOuter ? rxO : rxI;
		const ry = isOuter ? ryO : ryI;
		pts.push(`${cx + rx * Math.cos(ang)} ${cy + ry * Math.sin(ang)}`);
	}
	return `M ${pts[0]} ` + pts.slice(1).map(p => `L ${p}`).join(" ") + " Z";
}

// Wedge callout = rounded rectangle (optional corner radius) with a tail.
// adj1/adj2 = tail tip x,y in per-100000 of w,h (can be negative / >100000).
// adjR = corner radius fraction (we fold in as extra parameter so both the
// rect and roundRect callouts share this code).
function wedgeRectCallout(w: number, h: number, av: Adj, defaultCornerPer100k: number): string {
	const tipX = getAdj(av, "adj1", -20000) / 100000 * w;
	const tipY = getAdj(av, "adj2", 62500) / 100000 * h;
	const adj3 = getAdj(av, "adj3", defaultCornerPer100k);
	const r = Math.min((adj3 / 100000) * Math.min(w, h), Math.min(w, h) / 2);
	// Tail attaches to the nearest edge of the rect based on the angle from
	// the centre to the tip.
	const cx = w / 2, cy = h / 2;
	const dx = tipX - cx, dy = tipY - cy;
	let baseX1: number, baseY1: number, baseX2: number, baseY2: number;
	const tailHalf = Math.min(w, h) * 0.05;
	if (Math.abs(dx) / w > Math.abs(dy) / h) {
		// Tail emerges from left/right edge.
		const ex = dx < 0 ? 0 : w;
		baseX1 = ex; baseY1 = Math.max(r, Math.min(h - r, cy - tailHalf));
		baseX2 = ex; baseY2 = Math.max(r, Math.min(h - r, cy + tailHalf));
	} else {
		const ey = dy < 0 ? 0 : h;
		baseX1 = Math.max(r, Math.min(w - r, cx - tailHalf)); baseY1 = ey;
		baseX2 = Math.max(r, Math.min(w - r, cx + tailHalf)); baseY2 = ey;
	}
	// We'll just draw the (rounded) rect and then append the tail as a second
	// closed subpath — the fill-rule handles the overlap.
	const box = r > 0 ? roundRectPath(w, h, r) : rectPath(w, h);
	const tail = `M ${baseX1} ${baseY1} L ${tipX} ${tipY} L ${baseX2} ${baseY2} Z`;
	return `${box} ${tail}`;
}

function wedgeEllipseCallout(w: number, h: number, av: Adj): string {
	const tipX = getAdj(av, "adj1", -20000) / 100000 * w;
	const tipY = getAdj(av, "adj2", 62500) / 100000 * h;
	const cx = w / 2, cy = h / 2;
	const rx = w / 2, ry = h / 2;
	// Tail base: two points on the ellipse near the angle to the tip.
	const ang = Math.atan2((tipY - cy) / ry, (tipX - cx) / rx);
	const spread = 0.15; // radians
	const p1x = cx + rx * Math.cos(ang - spread);
	const p1y = cy + ry * Math.sin(ang - spread);
	const p2x = cx + rx * Math.cos(ang + spread);
	const p2y = cy + ry * Math.sin(ang + spread);
	const ell = ellipsePath(w, h);
	const tail = `M ${p1x} ${p1y} L ${tipX} ${tipY} L ${p2x} ${p2y} Z`;
	return `${ell} ${tail}`;
}
