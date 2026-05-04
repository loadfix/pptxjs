import type { ShapeLike } from '../presentation-parser';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';

// `hyperlinkUrls` maps relationship id → resolved URL for the owning slide.
// Passed down into shape/pic/text renderers so they can wrap runs and whole
// shapes with `<a>` anchors when `hyperlinkRId` is set on the parsed model.
export function renderShapeLike(shape: ShapeLike, cls: string, hyperlinkUrls: Map<string, string>): HTMLElement {
	if (shape.kind === 'pic') return renderPic(shape, cls, hyperlinkUrls);
	if (shape.kind === 'table') return renderTable(shape, cls, hyperlinkUrls);
	return renderShape(shape, cls, hyperlinkUrls);
}
