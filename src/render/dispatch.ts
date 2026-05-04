import type { ShapeLike } from '../presentation-parser';
import type { TableStyle } from '../table-style';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';
import { renderChartFallback, renderSmartArtFallback } from './chart-smartart';

// `hyperlinkUrls` maps relationship id → resolved URL for the owning slide.
// Passed down into shape/pic/text renderers so they can wrap runs and whole
// shapes with `<a>` anchors when `hyperlinkRId` is set on the parsed model.
export function renderShapeLike(
	shape: ShapeLike,
	cls: string,
	embedUrls: Map<string, string>,
	tableStyles: Map<string, TableStyle> | null,
	hyperlinkUrls: Map<string, string>,
): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls, hyperlinkUrls);
	if (shape.kind === 'table') return renderTable(shape, cls, tableStyles, hyperlinkUrls);
	if (shape.kind === 'chart-fallback') return renderChartFallback(shape, cls);
	if (shape.kind === 'smartart-fallback') return renderSmartArtFallback(shape, cls);
	return renderShape(shape, cls, embedUrls, hyperlinkUrls);
}
