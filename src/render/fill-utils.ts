// Extract a CSS color string from a Fill, for cases where we only know how
// to render solids today (shape borders / SVG strokes / backgrounds). Wave 2
// will extend this to render gradients/patterns/blips properly.

import type { Fill } from '../fill';

export function solidColorFromFill(fill: Fill | null): string | null {
	if (!fill) return null;
	if (fill.kind === 'solid') return fill.colorHex;
	return null;
}
