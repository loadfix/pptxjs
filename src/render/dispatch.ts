import type { ShapeLike } from '../presentation-parser';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';

export function renderShapeLike(shape: ShapeLike, cls: string, embedUrls: Map<string, string>): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls);
	if (shape.kind === 'table') return renderTable(shape, cls);
	return renderShape(shape, cls, embedUrls);
}
