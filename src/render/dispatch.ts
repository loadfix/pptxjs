import type { ShapeLike } from '../presentation-parser';
import type { TableStyle } from '../table-style';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';
import { renderChartFallback, renderSmartArtFallback } from './chart-smartart';
import type { FieldContext } from './text';

// `hyperlinkUrls` maps relationship id → resolved URL for the owning slide.
// Passed down into shape/pic/text renderers so they can wrap runs and whole
// shapes with `<a>` anchors when `hyperlinkRId` is set on the parsed model.
// `fieldCtx` carries the slide + firstSlideNum used for slidenum/ftr/hdr/dt
// field-run substitution and for suppressing hidden header/footer placeholders.
export function renderShapeLike(
	shape: ShapeLike,
	cls: string,
	embedUrls: Map<string, string>,
	tableStyles: Map<string, TableStyle> | null,
	hyperlinkUrls: Map<string, string>,
	fieldCtx?: FieldContext,
): HTMLElement | null {
	// Skip sldNum / ftr / hdr / dt placeholder shapes when the slide's
	// <p:hf> flags mark them hidden. Non-placeholder shapes, and placeholders
	// of other types (title, body, pic, etc.), are never suppressed here.
	if (fieldCtx && shape.kind === 'shape' && shape.phType) {
		const flags = fieldCtx.slide.hf;
		if (shape.phType === 'sldNum' && !flags.sldNum) return null;
		if (shape.phType === 'ftr' && !flags.ftr) return null;
		if (shape.phType === 'hdr' && !flags.hdr) return null;
		if (shape.phType === 'dt' && !flags.dt) return null;
	}
	if (shape.kind === 'pic') return renderPic(shape, cls, hyperlinkUrls);
	if (shape.kind === 'table') return renderTable(shape, cls, tableStyles, hyperlinkUrls, fieldCtx);
	if (shape.kind === 'chart-fallback') return renderChartFallback(shape, cls);
	if (shape.kind === 'smartart-fallback') return renderSmartArtFallback(shape, cls);
	return renderShape(shape, cls, embedUrls, hyperlinkUrls, fieldCtx);
}
