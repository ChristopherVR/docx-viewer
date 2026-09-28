import type { LayoutFontSpec, TextMeasurer } from './measure.js';
import { appendBreakMarker, tokenizeRun, type BreakToken } from './text-breaks.js';
import type { LayoutFragment, LayoutLine } from './result.js';
import type { LayoutParagraph, LayoutRun } from './input.js';
import { DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units.js';
import { placeTab } from './tab-stops.js';

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
}

function fontOf(run: LayoutRun): LayoutFontSpec {
	return {
		family: run.fontFamily || DEFAULT_FONT_FAMILY,
		sizePx: ptToPx(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT),
		bold: run.bold,
		italic: run.italic,
	};
}

function resolveIndents(paragraph: LayoutParagraph): {
	leftPx: number;
	rightPx: number;
	firstLineExtraPx: number;
} {
	const rtl = paragraph.direction === 'rtl';
	const startTwips =
		paragraph.indentStartTwips ??
		(rtl ? paragraph.indentRightTwips : paragraph.indentLeftTwips) ??
		0;
	const endTwips =
		paragraph.indentEndTwips ?? (rtl ? paragraph.indentLeftTwips : paragraph.indentRightTwips) ?? 0;
	const leftTwips = rtl ? endTwips : startTwips;
	const rightTwips = rtl ? startTwips : endTwips;
	const firstLineExtra =
		paragraph.hangingTwips != null ? -paragraph.hangingTwips : (paragraph.firstLineTwips ?? 0);
	return {
		leftPx: twipsToPx(leftTwips),
		rightPx: twipsToPx(rightTwips),
		firstLineExtraPx: twipsToPx(firstLineExtra),
	};
}

function tokenizeParagraph(paragraph: LayoutParagraph): {
	tokens: BreakToken[];
	runOffsets: number[];
} {
	const tokens: BreakToken[] = [];
	const runOffsets: number[] = [];
	let cursor = 0;
	paragraph.runs.forEach((run, runIndex) => {
		runOffsets.push(cursor);
		if (run.object) tokens.push({ kind: 'object', runIndex, sourceStart: 0 });
		else tokens.push(...tokenizeRun(run.text, runIndex));
		if (run.breakAfter) appendBreakMarker(tokens, runIndex, run.text.length, run.breakAfter);
		// A list label maps to the start of the paragraph's text for click-to-cursor.
		if (run.synthetic) {
			for (const token of tokens) if (token.runIndex === runIndex) token.sourceStart = 0;
		} else cursor += run.text.length;
	});
	return { tokens, runOffsets };
}

interface PlacedToken {
	token: BreakToken;
	widthPx: number;
	leader?: LayoutFragment['leader'];
}

function lineHeightForTokens(
	placed: PlacedToken[],
	fonts: LayoutFontSpec[],
	paragraph: LayoutParagraph,
	measurer: TextMeasurer,
): number {
	const usedFonts = placed.length
		? [...new Set(placed.map((p) => p.token.runIndex))].map((i) => fonts[i])
		: [fonts[0] ?? fontOf({ text: '' })];
	const textHeight = Math.max(...usedFonts.map((font) => measurer.lineHeightOf(font)), 0);
	// An inline picture raises its line to the picture's height (plus the text's descent share).
	const objectHeight = Math.max(
		0,
		...placed
			.filter((p) => p.token.kind === 'object')
			.map((p) => paragraph.runs[p.token.runIndex]?.object?.heightPx ?? 0),
	);
	const natural = Math.max(textHeight, objectHeight ? objectHeight + textHeight * 0.2 : 0);
	const rule = paragraph.lineSpacingRule ?? (paragraph.lineSpacingTwips != null ? 'auto' : 'auto');
	const spacingTwips = paragraph.lineSpacingTwips;
	if (rule === 'exact') return spacingTwips != null ? twipsToPx(spacingTwips) : natural;
	if (rule === 'atLeast')
		return Math.max(natural, spacingTwips != null ? twipsToPx(spacingTwips) : 0);
	const multiple = (spacingTwips ?? 240) / 240;
	return natural * multiple;
}

/**
 * Lays out one paragraph's lines within `contentWidthPx` (the column/table
 * cell width available to the paragraph, before indents are subtracted).
 * Page and column breaks are reported back as line indices; the page-flow
 * stage decides what to do with them (start a new page/column) since only it
 * knows the current page/column state.
 */
export function layoutParagraph(
	paragraph: LayoutParagraph,
	contentWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
): ParagraphLayoutResult {
	const { leftPx, rightPx, firstLineExtraPx } = resolveIndents(paragraph);
	const bodyWidth = Math.max(1, contentWidthPx - leftPx - rightPx);
	const fonts = paragraph.runs.map(fontOf);
	const { tokens, runOffsets } = tokenizeParagraph(paragraph);
	const globalOffset = (token: BreakToken) => runOffsets[token.runIndex] + token.sourceStart;
	const effectiveAlign = paragraph.align ?? (paragraph.direction === 'rtl' ? 'right' : 'left');

	const lines: LayoutLine[] = [];
	const pageBreakAfterLine = new Set<number>();
	const columnBreakAfterLine = new Set<number>();
	let widthOverflowNoted = false;

	let placed: PlacedToken[] = [];
	let lineWidthPx = 0;

	const availableWidth = (lineIndex: number) =>
		Math.max(1, bodyWidth - (lineIndex === 0 ? firstLineExtraPx : 0));
	/** Where a line's text starts, from the paragraph's text margin (indent plus first-line offset). */
	const lineStart = (lineIndex: number) => leftPx + (lineIndex === 0 ? firstLineExtraPx : 0);
	const tokenWidth = (token: BreakToken): number =>
		token.kind === 'word' || token.kind === 'space'
			? measurer.widthOf(token.text, fonts[token.runIndex])
			: token.kind === 'object'
				? (paragraph.runs[token.runIndex]?.object?.widthPx ?? 0)
				: 0;
	const widths = tokens.map(tokenWidth);
	const ends = new Set(['tab', 'lineBreak', 'pageBreak', 'columnBreak']);
	/** Width of the text after the tab at `index`, whole and up to its first decimal separator. */
	function segmentAfter(index: number): { followingPx: number; beforeDecimalPx: number } {
		let followingPx = 0;
		let beforeDecimalPx: number | undefined;
		for (let next = index + 1; next < tokens.length && !ends.has(tokens[next].kind); next++) {
			const token = tokens[next];
			if (beforeDecimalPx === undefined && token.kind === 'word' && /[.,]/.test(token.text)) {
				const prefix = token.text.slice(0, token.text.search(/[.,]/));
				beforeDecimalPx = followingPx + measurer.widthOf(prefix, fonts[token.runIndex]);
			}
			followingPx += widths[next];
		}
		return { followingPx, beforeDecimalPx: beforeDecimalPx ?? followingPx };
	}
	const hangingStopPx = firstLineExtraPx < 0 ? leftPx : undefined;

	function buildFragments(tokensOnLine: PlacedToken[]): LayoutFragment[] {
		const fragments: LayoutFragment[] = [];
		let x = 0;
		for (const { token, widthPx, leader } of tokensOnLine) {
			const run = paragraph.runs[token.runIndex];
			const text = token.kind === 'word' || token.kind === 'space' ? token.text : '';
			fragments.push({
				text,
				xPx: x,
				widthPx,
				runIndex: token.runIndex,
				bold: run?.bold,
				italic: run?.italic,
				fontFamily: run?.fontFamily,
				fontSizePt: run?.fontSizePt,
				...(leader ? { leader } : {}),
				...(run?.color ? { color: run.color } : {}),
				...(run?.underline ? { underline: true } : {}),
				...(run?.strike ? { strike: true } : {}),
				...(token.kind === 'object' && run?.object ? { object: run.object } : {}),
			});
			x += widthPx;
		}
		return fragments;
	}

	function justify(fragments: LayoutFragment[], width: number): LayoutFragment[] {
		const spaceIndices = fragments.map((f, i) => (f.text === ' ' ? i : -1)).filter((i) => i >= 0);
		if (!spaceIndices.length || !fragments.length) return fragments;
		const natural = fragments.at(-1)!.xPx + fragments.at(-1)!.widthPx;
		const extra = Math.max(0, width - natural) / spaceIndices.length;
		if (extra <= 0) return fragments;
		let shift = 0;
		return fragments.map((fragment, index) => {
			const placedFragment = { ...fragment, xPx: fragment.xPx + shift };
			if (spaceIndices.includes(index)) {
				placedFragment.widthPx += extra;
				shift += extra;
			}
			return placedFragment;
		});
	}

	function alignFragments(
		fragments: LayoutFragment[],
		width: number,
		isLastLine: boolean,
	): LayoutFragment[] {
		if (!fragments.length) return fragments;
		if (effectiveAlign === 'justify' && !isLastLine) return justify(fragments, width);
		const natural = fragments.at(-1)!.xPx + fragments.at(-1)!.widthPx;
		const offset =
			effectiveAlign === 'center'
				? Math.max(0, width - natural) / 2
				: effectiveAlign === 'right'
					? Math.max(0, width - natural)
					: 0;
		return offset === 0 ? fragments : fragments.map((f) => ({ ...f, xPx: f.xPx + offset }));
	}

	function flushLine(forceEmpty: boolean, forcedByBreak: 'page' | 'column' | undefined) {
		if (!placed.length && !forceEmpty) return;
		const lineIndex = lines.length;
		// Trailing whitespace at a wrap point is consumed (counts toward sourceEnd)
		// but never rendered or measured, so it cannot skew alignment/justification.
		const rendered = [...placed];
		while (rendered.length && rendered.at(-1)!.token.kind === 'space') rendered.pop();
		const fragments = buildFragments(rendered);
		const start = placed.length ? globalOffset(placed[0].token) : 0;
		const last = placed.at(-1)?.token;
		const end = last
			? globalOffset(last) + (last.kind === 'word' || last.kind === 'space' ? last.text.length : 1)
			: start;
		lines.push({
			yPx: 0,
			heightPx: lineHeightForTokens(placed, fonts, paragraph, measurer),
			fragments,
			sourceStart: start,
			sourceEnd: end,
		});
		if (forcedByBreak === 'page') pageBreakAfterLine.add(lineIndex);
		if (forcedByBreak === 'column') columnBreakAfterLine.add(lineIndex);
		placed = [];
		lineWidthPx = 0;
	}

	for (const [tokenIndex, token] of tokens.entries()) {
		if (token.kind === 'lineBreak') {
			flushLine(true, undefined);
			continue;
		}
		if (token.kind === 'pageBreak' || token.kind === 'columnBreak') {
			flushLine(true, token.kind === 'pageBreak' ? 'page' : 'column');
			continue;
		}
		let leader: LayoutFragment['leader'];
		let width = widths[tokenIndex];
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
			flushLine(false, undefined);
			if (width > availableWidth(lines.length) && !widthOverflowNoted) {
				widthOverflowNoted = true;
				note(
					'A word wider than its line/column was not split mid-word; it overflows the line box.',
				);
			}
		}
		if (token.kind === 'space' && placed.length === 0) continue; // leading space at wrap point is dropped
		placed.push({ token, widthPx: width, ...(leader ? { leader } : {}) });
		lineWidthPx += width;
	}
	flushLine(lines.length === 0, undefined);

	const lastIndex = lines.length - 1;
	for (let index = 0; index < lines.length; index++) {
		lines[index] = {
			...lines[index],
			fragments: alignFragments(
				lines[index].fragments,
				availableWidth(index),
				index === lastIndex,
			).map((fragment) => ({ ...fragment, xPx: fragment.xPx + lineStart(index) })),
		};
	}

	let y = 0;
	for (const line of lines) {
		line.yPx = y;
		y += line.heightPx;
	}

	return {
		lines,
		spacingBeforePx: twipsToPx(paragraph.spacingBeforeTwips ?? 0),
		spacingAfterPx: twipsToPx(paragraph.spacingAfterTwips ?? 0),
		contentHeightPx: y,
		pageBreakAfterLine,
		columnBreakAfterLine,
	};
}
