import type { ShapeLike } from '../presentation-parser';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';
import { renderChartFallback, renderSmartArtFallback } from './chart-smartart';

export function renderShapeLike(shape: ShapeLike, cls: string): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls);
	if (shape.kind === 'table') return renderTable(shape, cls);
	if (shape.kind === 'chart-fallback') return renderChartFallback(shape, cls);
	if (shape.kind === 'smartart-fallback') return renderSmartArtFallback(shape, cls);
	return renderShape(shape, cls);
}
