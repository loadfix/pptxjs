import type { ShapeLike } from '../presentation-parser';
import { renderShape } from './shape';
import { renderPic } from './pic';
import { renderTable } from './table';
import type { FieldContext } from './text';

export function renderShapeLike(shape: ShapeLike, cls: string, fieldCtx?: FieldContext): HTMLElement | null {
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
	if (shape.kind === 'pic') return renderPic(shape, cls);
	if (shape.kind === 'table') return renderTable(shape, cls, fieldCtx);
	return renderShape(shape, cls, fieldCtx);
}
