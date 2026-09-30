import { definedProps } from './defined-props.js';
import type { LayoutFontSpec, TextMeasurer } from './measure.js';
import type { LayoutFragment } from './result.js';
import { appendBreakMarker, tokenizeRun, type BreakToken } from './text-breaks.js';
import type { LayoutParagraph, LayoutRun } from './input.js';
import { NO_TWIPS, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units.js';

/** Word draws superscript and subscript text at about two thirds of the run's size. */
export const SCRIPT_SCALE = 0.65;

export function fontOf(run: LayoutRun): LayoutFontSpec {
	return {
		family: run.fontFamily || DEFAULT_FONT_FAMILY,
		sizePx: ptToPx(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (run.script ? SCRIPT_SCALE : 1),
		...definedProps({ bold: run.bold, italic: run.italic, ligatures: run.ligatures }),
		...(run.kerningThresholdPt === undefined
			? {}
			: {
					kerning:
						run.kerningThresholdPt > 0 &&
						(run.fontSizePt ?? DEFAULT_FONT_SIZE_PT) >= run.kerningThresholdPt
							? ('normal' as const)
							: ('none' as const),
				}),
	};
}

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
/** Word applies additional character spacing after scaling the glyph advances. */
export function runTextWidth(text: string, run: LayoutRun, measurer: TextMeasurer): number {
	return Math.max(
		0,
		(measurer.widthOf(text, fontOf(run)) * (run.textScalePercent ?? 100)) / 100 +
			[...graphemes.segment(text)].length * (run.characterSpacingPx ?? 0),
	);
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
		NO_TWIPS;
	const endTwips =
		paragraph.indentEndTwips ??
		(rtl ? paragraph.indentLeftTwips : paragraph.indentRightTwips) ??
		NO_TWIPS;
	const leftTwips = rtl ? endTwips : startTwips;
	const rightTwips = rtl ? startTwips : endTwips;
	const firstLineExtraPx =
		paragraph.hangingTwips != null
			? -twipsToPx(paragraph.hangingTwips)
			: twipsToPx(paragraph.firstLineTwips ?? NO_TWIPS);
	return {
		leftPx: twipsToPx(leftTwips),
		rightPx: twipsToPx(rightTwips),
		firstLineExtraPx,
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
		else if (run.marker) {
			// A marker such as "Chapter 1.2" is indivisible; its separator still acts
			// as a real tab or space after the aligned glyphs.
			tokens.push({
				kind: 'word',
				runIndex,
				sourceStart: 0,
				text: run.text.slice(0, run.marker.length),
			});
			tokens.push(...tokenizeRun(run.text.slice(run.marker.length), runIndex));
		} else tokens.push(...tokenizeRun(run.text, runIndex));
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
	/** Glyph position and consumed advance differ for centered/right-aligned list markers. */
	xOffsetPx?: number;
	advancePx?: number;
	leader?: LayoutFragment['leader'];
}
