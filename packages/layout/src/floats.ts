import type { LayoutDocumentInput, LayoutFloat } from './input.js';
import type { LayoutColumnBox, LayoutFloatBox, LayoutPageBox } from './result.js';

export const FLOAT_WRAP_NOTE =
	'Floating pictures are drawn at their anchored positions, but body text does not wrap around them in Print Layout.';

interface Span {
	start: number;
	size: number;
}

function horizontalArea(
	float: LayoutFloat,
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
): Span {
	const contentLeft = page.marginLeftPx;
	switch (float.relativeFromH) {
		case 'page':
			return { start: 0, size: page.widthPx };
		case 'margin':
			return { start: contentLeft, size: page.widthPx - contentLeft - page.marginRightPx };
		case 'leftMargin':
		case 'insideMargin':
			return { start: 0, size: contentLeft };
		case 'rightMargin':
		case 'outsideMargin':
			return { start: page.widthPx - page.marginRightPx, size: page.marginRightPx };
		default:
			// `column` (Word's default) and `character` (approximated by the column).
			return { start: contentLeft + column.xPx, size: column.widthPx };
	}
}

function verticalArea(
	float: LayoutFloat,
	page: LayoutPageBox,
	paragraphTopPx: number,
	paragraphHeightPx: number,
): Span {
	switch (float.relativeFromV) {
		case 'page':
			return { start: 0, size: page.heightPx };
		case 'margin':
			return {
				start: page.marginTopPx,
				size: page.heightPx - page.marginTopPx - page.marginBottomPx,
			};
		case 'topMargin':
		case 'insideMargin':
			return { start: 0, size: page.marginTopPx };
		case 'bottomMargin':
		case 'outsideMargin':
			return { start: page.heightPx - page.marginBottomPx, size: page.marginBottomPx };
		default:
			// `paragraph` (Word's default) and `line` (approximated by the paragraph's first line).
			return { start: paragraphTopPx, size: paragraphHeightPx };
	}
}

function place(
	area: Span,
	size: number,
	align: string | undefined,
	offset: number | undefined,
): number {
	if (align === 'center') return area.start + (area.size - size) / 2;
	if (align === 'right' || align === 'bottom' || align === 'outside')
		return area.start + area.size - size;
	if (align) return area.start;
	return area.start + (offset ?? 0);
}

/**
 * A floating picture's top-left on the page (CSS pixels), from its horizontal and vertical frame,
 * alignment and offset. `paragraph` is the anchor paragraph's page-relative top and height.
 */
export function floatPosition(
	float: LayoutFloat,
	page: LayoutPageBox,
	column: Pick<LayoutColumnBox, 'xPx' | 'widthPx'>,
	paragraph: { topPx: number; heightPx: number },
): { xPx: number; yPx: number } {
	const h = horizontalArea(float, page, column);
	const v = verticalArea(float, page, paragraph.topPx, paragraph.heightPx);
	return {
		xPx: place(h, float.widthPx, float.alignH, float.offsetXPx),
		yPx: place(v, float.heightPx, float.alignV, float.offsetYPx),
	};
}

/**
 * Positions each paragraph's floating pictures on the page where the paragraph starts, from their
 * `wp:positionH`/`wp:positionV` reference frame, alignment and offset. Returns whether any were placed.
 */
export function positionFloats(input: LayoutDocumentInput, pages: LayoutPageBox[]): boolean {
	const floatsById = new Map<string, LayoutFloat[]>();
	for (const section of input.sections)
		for (const block of section.blocks)
			if (block.kind === 'paragraph' && block.floats?.length)
				floatsById.set(block.id, block.floats);
	if (!floatsById.size) return false;
	const placed = new Set<string>();
	for (const page of pages) {
		const boxes: LayoutFloatBox[] = [];
		for (const column of page.columns)
			for (const block of column.blocks) {
				const floats = block.kind === 'paragraph' ? floatsById.get(block.blockId) : undefined;
				if (!floats || placed.has(block.blockId)) continue;
				placed.add(block.blockId);
				const top = page.marginTopPx + block.yPx;
				for (const float of floats) {
					const position = floatPosition(float, page, column, {
						topPx: top,
						heightPx: block.heightPx,
					});
					boxes.push({
						blockId: block.blockId,
						...position,
						widthPx: float.widthPx,
						heightPx: float.heightPx,
						partName: float.partName,
						contentType: float.contentType,
						behindText: Boolean(float.behindText),
					});
				}
			}
		if (boxes.length) page.floats = boxes;
	}
	return true;
}
