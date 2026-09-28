import { definedProps } from './defined-props.js';
import type { LayoutFontSpec } from './measure.js';
import type { LayoutFragment } from './result.js';
import { appendBreakMarker, tokenizeRun, type BreakToken } from './text-breaks.js';
import type { LayoutParagraph, LayoutRun } from './input.js';
import { DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units.js';

/** Word draws superscript and subscript text at about two thirds of the run's size. */
export const SCRIPT_SCALE = 0.65;

export function fontOf(run: LayoutRun): LayoutFontSpec {
	return {
		family: run.fontFamily || DEFAULT_FONT_FAMILY,
		sizePx: ptToPx(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (run.script ? SCRIPT_SCALE : 1),
		...definedProps({ bold: run.bold, italic: run.italic }),
	};
}

export function resolveIndents(paragraph: LayoutParagraph): {
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

export function tokenizeParagraph(paragraph: LayoutParagraph): {
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

export interface PlacedToken {
	token: BreakToken;
	widthPx: number;
	leader?: LayoutFragment['leader'];
}
