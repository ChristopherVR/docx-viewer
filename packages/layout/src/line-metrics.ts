import { fontMetrics, type LayoutFontSpec, type TextMeasurer } from './measure.js';
import type { BreakToken } from './text-breaks.js';
import type { LayoutParagraph } from './input.js';
import { DEFAULT_FONT_SIZE_PT, ptToPx, twipsToPx } from './units.js';
import { fontOf, type PlacedToken } from './paragraph-tokens.js';

/** How far superscript text is raised and subscript lowered, as a share of the run's size. */
const SUPER_RAISE = 0.33;
const SUB_DROP = 0.14;

/** Where a token's box sits relative to the baseline: its extent above and below it. */
export function tokenExtent(
	token: BreakToken,
	paragraph: LayoutParagraph,
	fonts: LayoutFontSpec[],
	measurer: TextMeasurer,
): { above: number; below: number; ascent: number; descent: number } {
	const run = paragraph.runs[token.runIndex];
	if (token.kind === 'object' && run?.object)
		return { above: run.object.heightPx, below: 0, ascent: run.object.heightPx, descent: 0 };
	const { ascent, descent } = fontMetrics(measurer, fonts[token.runIndex] ?? fontOf({ text: '' }));
	const size = ptToPx(run?.fontSizePt ?? DEFAULT_FONT_SIZE_PT);
	const shift =
		run?.script === 'super' ? SUPER_RAISE * size : run?.script === 'sub' ? -SUB_DROP * size : 0;
	return { above: ascent + shift, below: descent - shift, ascent, descent };
}

/**
 * A line's height and baseline: the tallest extent above and below the baseline over its tokens
 * (inline pictures sit on the baseline), scaled by the paragraph's line spacing. Word adds extra
 * spacing above the text, so the baseline sits one descent above the line's bottom.
 */
export function lineMetrics(
	placed: PlacedToken[],
	fonts: LayoutFontSpec[],
	paragraph: LayoutParagraph,
	measurer: TextMeasurer,
): { heightPx: number; baselinePx: number } {
	const tokens = placed.length
		? placed.map((p) => p.token)
		: [{ kind: 'space', text: ' ', runIndex: 0, sourceStart: 0 } as BreakToken];
	let above = 0;
	let below = 0;
	for (const token of tokens) {
		const extent = tokenExtent(token, paragraph, fonts, measurer);
		above = Math.max(above, extent.above);
		below = Math.max(below, extent.below);
	}
	const natural = above + below;
	const rule = paragraph.lineSpacingRule ?? 'auto';
	const spacingTwips = paragraph.lineSpacingTwips;
	const heightPx =
		rule === 'exact'
			? spacingTwips != null
				? twipsToPx(spacingTwips)
				: natural
			: rule === 'atLeast'
				? Math.max(natural, spacingTwips != null ? twipsToPx(spacingTwips) : 0)
				: natural * ((spacingTwips ?? 240) / 240);
	return { heightPx, baselinePx: heightPx - below };
}
