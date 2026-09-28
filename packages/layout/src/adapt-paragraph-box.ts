import type { DocumentModel, ParagraphFormatting } from '@christophervr/docx-core';
import { border, cssHex } from './adapt-table.js';
import type { LayoutBlock, LayoutParagraph, LayoutParagraphBorders } from './input.js';

const pointsToPx = (points: number) => (points * 96) / 72;

/** Paragraph borders and shading from resolved formatting. */
export function paragraphBox(
	formatting: ParagraphFormatting,
	theme: DocumentModel['theme'],
): Pick<LayoutParagraph, 'borders' | 'shading'> & {
	betweenBorder?: LayoutParagraphBorders['top'];
} {
	const borders: LayoutParagraphBorders = {};
	let betweenBorder: LayoutParagraphBorders['top'];
	for (const side of ['top', 'bottom', 'left', 'right', 'between'] as const) {
		const source = formatting.borders?.[side];
		const line = border(source, theme);
		if (!line) continue;
		const resolved = { ...line, spacePx: pointsToPx(source?.spacePoints ?? 0) };
		if (side === 'between') betweenBorder = resolved;
		else borders[side] = resolved;
	}
	const shading = cssHex(formatting.shadingFill);
	return {
		...(Object.keys(borders).length ? { borders } : {}),
		...(shading ? { shading } : {}),
		...(betweenBorder ? { betweenBorder } : {}),
	};
}

/**
 * Word draws consecutive paragraphs with identical borders as one box: the top line only above
 * the first, the bottom line only below the last, and the `between` line (if any) between them.
 */
export function groupParagraphBorders(
	blocks: LayoutBlock[],
	between: Map<LayoutParagraph, LayoutParagraphBorders['top']>,
): void {
	// Keys come from the original borders, before any paragraph in the group is adjusted.
	const keys = blocks.map((block) =>
		block.kind === 'paragraph' && block.borders ? JSON.stringify(block.borders) : undefined,
	);
	blocks.forEach((block, index) => {
		if (block.kind !== 'paragraph' || !block.borders) return;
		const sameAsPrevious = keys[index - 1] === keys[index];
		const sameAsNext = keys[index + 1] === keys[index];
		if (!sameAsPrevious && !sameAsNext) return;
		const borders = { ...block.borders };
		if (sameAsNext) delete borders.bottom;
		if (sameAsPrevious) {
			const line = between.get(block);
			if (line) borders.top = line;
			else delete borders.top;
		}
		block.borders = borders;
	});
}
