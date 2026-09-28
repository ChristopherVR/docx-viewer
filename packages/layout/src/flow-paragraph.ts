import type { PageCursor } from './page-cursor.js';
import { adjustForWidowOrphan, widowControlEnabled } from './keep-rules.js';
import type { ParagraphLayoutResult } from './paragraph-layout.js';
import type { LayoutParagraph } from './input.js';
import type { LayoutParagraphBox } from './result.js';
import { expectDefined } from './expect-defined.js';

export interface ParagraphPlacement {
	paragraph: LayoutParagraph;
	layout: ParagraphLayoutResult;
	/** Spacing-before in px, with `w:contextualSpacing` suppression already resolved by the caller. */
	spacingBeforePx: number;
	/**
	 * Minimum height (including this paragraph and any `keepNext` chain after
	 * it, as resolved by the caller) that must fit before this paragraph may
	 * start; 0 when neither `keepNext` nor `keepLines` apply.
	 */
	requiredTogetherPx: number;
}

/**
 * Places one paragraph's lines into `cursor`, splitting across page/column
 * boundaries as needed. Honors `pageBreakBefore`, the keep-together height
 * computed by the caller (chained `keepNext`/whole-paragraph `keepLines`),
 * default 2-line widow/orphan control, and any explicit page/column break
 * tokens carried on the paragraph's runs.
 */
export function placeParagraph(
	cursor: PageCursor,
	placement: ParagraphPlacement,
	note: (message: string) => void,
): void {
	const { paragraph, layout, spacingBeforePx, requiredTogetherPx } = placement;
	if (paragraph.pageBreakBefore && !cursor.atColumnTop) cursor.newPage();
	if (
		requiredTogetherPx > 0 &&
		requiredTogetherPx <= cursor.columnHeightPx &&
		requiredTogetherPx > cursor.remainingHeightPx() &&
		!cursor.atColumnTop
	) {
		cursor.newColumn();
	}

	const widow = widowControlEnabled(paragraph.widowControl);
	let lineCursor = 0;
	let firstSegment = true;
	if (!layout.lines.length) {
		// An empty paragraph still occupies its spacing/blank-line height.
		cursor.place(
			{ kind: 'paragraph', blockId: paragraph.id, yPx: 0, heightPx: 0, lines: [] },
			spacingBeforePx + layout.spacingAfterPx,
		);
		return;
	}
	while (lineCursor < layout.lines.length) {
		const budget = cursor.remainingHeightPx() - (firstSegment ? spacingBeforePx : 0);
		let cumulative = 0;
		let end = lineCursor;
		let forced: 'page' | 'column' | undefined;
		for (let i = lineCursor; i < layout.lines.length; i++) {
			// A line's gap (clearing a wrapped picture) moves with it, except when a later segment
			// starts it at the top of a new column or page.
			const line = expectDefined(layout.lines[i], 'paragraph line index');
			const gap = i === lineCursor && !firstSegment ? 0 : (line.gapBeforePx ?? 0);
			const height = line.heightPx + gap;
			// A line that does not fit waits for the next column, unless the column is empty.
			if (cumulative + height > budget && (end > lineCursor || !cursor.atColumnTop)) break;
			cumulative += height;
			end = i + 1;
			if (layout.pageBreakAfterLine.has(i)) {
				forced = 'page';
				break;
			}
			if (layout.columnBreakAfterLine.has(i)) {
				forced = 'column';
				break;
			}
		}
		let placedCount = end - lineCursor;
		if (!forced && end < layout.lines.length)
			placedCount = adjustForWidowOrphan(layout.lines.length - lineCursor, placedCount, widow);
		if (placedCount === 0) {
			if (!cursor.atColumnTop) {
				cursor.newColumn();
				continue;
			}
			note(
				'A paragraph line was taller than an empty page/column and was placed without splitting.',
			);
			placedCount = Math.max(1, end - lineCursor);
		}
		const first = expectDefined(layout.lines[lineCursor], 'first line of paragraph segment');
		const origin = first.yPx - (firstSegment ? (first.gapBeforePx ?? 0) + layout.insetTopPx : 0);
		const segmentLines = layout.lines
			.slice(lineCursor, lineCursor + placedCount)
			.map((line) => ({ ...line, yPx: line.yPx - origin }));
		const isFinal = lineCursor + placedCount >= layout.lines.length;
		const lastLine = segmentLines.at(-1);
		const isLastSegment = lineCursor + placedCount >= layout.lines.length;
		const segmentHeight =
			(lastLine ? lastLine.yPx + lastLine.heightPx : 0) +
			(isLastSegment ? layout.insetBottomPx : 0);
		const box: LayoutParagraphBox = {
			kind: 'paragraph',
			blockId: paragraph.id,
			yPx: 0,
			heightPx: segmentHeight,
			lines: segmentLines,
			...(layout.frame ? { frame: layout.frame } : {}),
		};
		const advance =
			(firstSegment ? spacingBeforePx : 0) + segmentHeight + (isFinal ? layout.spacingAfterPx : 0);
		cursor.place(box, advance);
		lineCursor += placedCount;
		firstSegment = false;
		if (forced === 'page') cursor.newPage();
		else if (forced === 'column') cursor.newColumn();
		else if (lineCursor < layout.lines.length) cursor.newColumn();
	}
}
