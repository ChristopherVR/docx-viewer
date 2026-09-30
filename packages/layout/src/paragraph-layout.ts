import type { TextMeasurer } from './measure.js';
import type { BreakToken } from './text-breaks.js';
import type { LayoutFragment, LayoutLine, LayoutParagraphFrame } from './result.js';
import type { LayoutParagraph } from './input.js';
import { NO_TWIPS, twipsToPx } from './units.js';
import { placeTab } from './tab-stops.js';
import {
	fontOf,
	runTextWidth,
	resolveIndents,
	tokenizeParagraph,
	type PlacedToken,
} from './paragraph-tokens.js';
import { lineMetrics, tokenExtent } from './line-metrics.js';
import { alignFragments, buildFragments } from './paragraph-fragments.js';
import { expectDefined } from './expect-defined.js';

export interface ParagraphLayoutResult {
	/** Lines with `yPx` relative to the paragraph box's own top (0 for the first line). */
	lines: LayoutLine[];
	spacingBeforePx: number;
	spacingAfterPx: number;
	contentHeightPx: number;
	/** Indices (into `lines`) after which an explicit page break token was consumed. */
	pageBreakAfterLine: Set<number>;
	/** Indices after which an explicit column break token was consumed. */
	columnBreakAfterLine: Set<number>;
	/** Space above the first line and below the last taken by top and bottom borders. */
	insetTopPx: number;
	insetBottomPx: number;
	frame?: LayoutParagraphFrame;
}

/**
 * Lays out one paragraph's lines within `contentWidthPx` (the column/table
 * cell width available to the paragraph, before indents are subtracted).
 * Page and column breaks are reported back as line indices; the page-flow
 * stage decides what to do with them (start a new page/column) since only it
 * knows the current page/column state.
 */
/**
 * Space taken from one line by wrapped floating pictures: insets from the line's left and right
 * (beyond the paragraph indents) and a gap to skip first (text below a picture). `yPx` is the
 * line's top relative to the paragraph's first line.
 */
export type LineBoxFn = (
	yPx: number,
	heightPx: number,
) => { leftInsetPx: number; rightInsetPx: number; gapBeforePx: number };

export function layoutParagraph(
	paragraph: LayoutParagraph,
	contentWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	lineBox?: LineBoxFn,
	startToken = 0,
): ParagraphLayoutResult {
	const { leftPx, rightPx, firstLineExtraPx: initialIndent } = resolveIndents(paragraph);
	const firstLineExtraPx = startToken ? 0 : initialIndent;
	const bodyWidth = Math.max(1, contentWidthPx - leftPx - rightPx);
	const fonts = paragraph.runs.map(fontOf);
	const { tokens, runOffsets } = tokenizeParagraph(paragraph);
	const globalOffset = (token: BreakToken) =>
		expectDefined(runOffsets[token.runIndex], 'run offset index') + token.sourceStart;
	const effectiveAlign = paragraph.align ?? (paragraph.direction === 'rtl' ? 'right' : 'left');

	const lines: LayoutLine[] = [];
	const pageBreakAfterLine = new Set<number>();
	const columnBreakAfterLine = new Set<number>();
	let widthOverflowNoted = false;

	let placed: PlacedToken[] = [];
	let lineWidthPx = 0;
	let tokenCursor = startToken;
	let firstToken = startToken;

	// Per-line boxes around wrapped floats, computed when each line starts (its top is then known).
	const boxes: { leftInsetPx: number; rightInsetPx: number; gapBeforePx: number }[] = [];
	let runningY = 0;
	const guessHeight = measurer.lineHeightOf(fonts[0] ?? fontOf({ text: '' }));
	const boxFor = (lineIndex: number) => {
		if (!boxes[lineIndex]) {
			let gap = 0;
			let box = lineBox?.(runningY, guessHeight);
			// Skipping below one picture can reach another; settle within a few steps.
			for (let step = 0; box && box.gapBeforePx > 0 && step < 8; step++) {
				gap += box.gapBeforePx;
				box = lineBox!(runningY + gap, guessHeight);
			}
			boxes[lineIndex] = {
				leftInsetPx: box?.leftInsetPx ?? 0,
				rightInsetPx: box?.rightInsetPx ?? 0,
				gapBeforePx: gap,
			};
		}
		return expectDefined(boxes[lineIndex], 'line box');
	};
	const availableWidth = (lineIndex: number) => {
		const box = boxFor(lineIndex);
		return Math.max(
			1,
			bodyWidth - (lineIndex === 0 ? firstLineExtraPx : 0) - box.leftInsetPx - box.rightInsetPx,
		);
	};
	/** Where a line's text starts, from the paragraph's text margin (indent plus first-line offset). */
	const lineStart = (lineIndex: number) =>
		leftPx + (lineIndex === 0 ? firstLineExtraPx : 0) + boxFor(lineIndex).leftInsetPx;
	const tokenWidth = (token: BreakToken): number =>
		token.kind === 'word' || token.kind === 'space'
			? runTextWidth(
					token.text,
					expectDefined(paragraph.runs[token.runIndex], 'token run'),
					measurer,
				)
			: token.kind === 'object'
				? (paragraph.runs[token.runIndex]?.object?.widthPx ?? 0)
				: 0;
	const widths = tokens.map(tokenWidth);
	const ends = new Set(['tab', 'lineBreak', 'pageBreak', 'columnBreak']);
	/** Width of the text after the tab at `index`, whole and up to its first decimal separator. */
	function segmentAfter(index: number): { followingPx: number; beforeDecimalPx: number } {
		let followingPx = 0;
		let beforeDecimalPx: number | undefined;
		for (let next = index + 1; next < tokens.length; next++) {
			const token = expectDefined(tokens[next], 'token index');
			if (ends.has(token.kind)) break;
			if (beforeDecimalPx === undefined && token.kind === 'word' && /[.,]/.test(token.text)) {
				const prefix = token.text.slice(0, token.text.search(/[.,]/));
				beforeDecimalPx =
					followingPx +
					runTextWidth(
						prefix,
						expectDefined(paragraph.runs[token.runIndex], 'decimal run'),
						measurer,
					);
			}
			followingPx += widths[next] ?? 0;
		}
		return { followingPx, beforeDecimalPx: beforeDecimalPx ?? followingPx };
	}
	const hangingStopPx = firstLineExtraPx < 0 ? leftPx : undefined;

	function flushLine(forceEmpty: boolean, forcedByBreak: 'page' | 'column' | undefined) {
		if (!placed.length && !forceEmpty) return;
		const lineIndex = lines.length;
		// Trailing whitespace at a wrap point is consumed (counts toward sourceEnd)
		// but never rendered or measured, so it cannot skew alignment/justification.
		const rendered = [...placed];
		while (rendered.length && rendered.at(-1)!.token.kind === 'space') rendered.pop();
		const fragments = buildFragments(rendered, paragraph);
		const [firstPlaced] = placed;
		const start = firstPlaced ? globalOffset(firstPlaced.token) : 0;
		const last = placed.at(-1)?.token;
		const end = last
			? globalOffset(last) + (last.kind === 'word' || last.kind === 'space' ? last.text.length : 1)
			: start;
		const gapBeforePx = boxFor(lineIndex).gapBeforePx;
		const { heightPx, baselinePx } = lineMetrics(placed, fonts, paragraph, measurer);
		// Each fragment's box is placed so its baseline sits on the line's.
		rendered.forEach(({ token }, index) => {
			const extent = tokenExtent(token, paragraph, fonts, measurer);
			const fragment = expectDefined(fragments[index], 'fragment for rendered token');
			fragment.topPx = baselinePx - extent.above;
			if (token.kind !== 'object') fragment.boxHeightPx = extent.ascent + extent.descent;
		});
		lines.push({
			yPx: 0,
			heightPx,
			baselinePx,
			fragments,
			sourceStart: start,
			sourceEnd: end,
			nextToken: tokenCursor,
			firstToken,
			...(gapBeforePx ? { gapBeforePx } : {}),
		});
		runningY += gapBeforePx + heightPx;
		if (forcedByBreak === 'page') pageBreakAfterLine.add(lineIndex);
		if (forcedByBreak === 'column') columnBreakAfterLine.add(lineIndex);
		placed = [];
		lineWidthPx = 0;
	}

	for (const [tokenIndex, token] of tokens.entries()) {
		if (tokenIndex < startToken) continue;
		tokenCursor = tokenIndex + 1;
		if (token.kind === 'lineBreak') {
			if (!placed.length) firstToken = tokenIndex;
			flushLine(true, undefined);
			continue;
		}
		if (token.kind === 'pageBreak' || token.kind === 'columnBreak') {
			if (!placed.length) firstToken = tokenIndex;
			flushLine(true, token.kind === 'pageBreak' ? 'page' : 'column');
			continue;
		}
		let leader: LayoutFragment['leader'];
		let width = expectDefined(widths[tokenIndex], 'token width');
		if (token.kind === 'tab') {
			const { followingPx, beforeDecimalPx } = segmentAfter(tokenIndex);
			const tab = placeTab(
				lineStart(lines.length) + lineWidthPx,
				paragraph.tabStops ?? [],
				hangingStopPx,
				followingPx,
				beforeDecimalPx,
			);
			width = tab.widthPx;
			leader = tab.leader === 'none' ? undefined : tab.leader;
		}
		const lineIndex = lines.length;
		const limit = availableWidth(lineIndex);
		// A pending trailing space only ever gets counted once it turns out to be
		// internal (i.e. this token fits after it); a trailing space at an actual
		// wrap point is trimmed in flushLine instead, so it must not be
		// pre-subtracted here.
		const wouldFit = lineWidthPx + width <= limit || placed.length === 0;
		if (!wouldFit && token.kind !== 'space') {
			tokenCursor = tokenIndex;
			flushLine(false, undefined);
			tokenCursor = tokenIndex + 1;
			if (width > availableWidth(lines.length) && !widthOverflowNoted) {
				widthOverflowNoted = true;
				note(
					'A word wider than its line/column was not split mid-word; it overflows the line box.',
				);
			}
		}
		if (token.kind === 'space' && placed.length === 0) continue; // leading space at wrap point is dropped
		if (!placed.length) firstToken = tokenIndex;
		placed.push({ token, widthPx: width, ...(leader ? { leader } : {}) });
		lineWidthPx += width;
	}
	flushLine(lines.length === 0, undefined);

	const lastIndex = lines.length - 1;
	for (const [index, line] of lines.entries()) {
		lines[index] = {
			...line,
			fragments: alignFragments(
				line.fragments,
				availableWidth(index),
				index === lastIndex,
				effectiveAlign,
			).map((fragment) => ({ ...fragment, xPx: fragment.xPx + lineStart(index) })),
		};
	}

	// Borders sit outside the text: top and bottom lines (with their gaps) add to the height, and
	// side lines are drawn beyond the indents.
	const borders = paragraph.borders;
	const extent = (side: keyof NonNullable<typeof borders>) =>
		borders?.[side] ? borders[side]!.widthPx + borders[side]!.spacePx : 0;
	const insetTopPx = extent('top');
	const insetBottomPx = extent('bottom');
	let y = insetTopPx;
	for (const line of lines) {
		y += line.gapBeforePx ?? 0;
		line.yPx = y;
		y += line.heightPx;
	}
	const frame =
		borders || paragraph.shading
			? {
					leftPx: leftPx - extent('left'),
					widthPx: bodyWidth + extent('left') + extent('right'),
					...(borders ? { borders } : {}),
					...(paragraph.shading ? { shading: paragraph.shading } : {}),
				}
			: undefined;

	return {
		lines,
		spacingBeforePx: twipsToPx(paragraph.spacingBeforeTwips ?? NO_TWIPS),
		spacingAfterPx: twipsToPx(paragraph.spacingAfterTwips ?? NO_TWIPS),
		contentHeightPx: y + insetBottomPx,
		insetTopPx,
		insetBottomPx,
		...(frame ? { frame } : {}),
		pageBreakAfterLine,
		columnBreakAfterLine,
	};
}
