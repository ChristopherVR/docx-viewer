import { definedProps } from './defined-props.js';
import type { LayoutFragment } from './result.js';
import type { LayoutParagraph } from './input.js';
import type { PlacedToken } from './paragraph-tokens.js';

export function buildFragments(
	tokensOnLine: PlacedToken[],
	paragraph: LayoutParagraph,
): LayoutFragment[] {
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
			...definedProps({
				bold: run?.bold,
				italic: run?.italic,
				fontFamily: run?.fontFamily,
				fontSizePt: run?.fontSizePt,
				textScalePercent: run?.textScalePercent,
				ligatures: run?.ligatures,
				characterSpacingPx: run?.characterSpacingPx,
				kerningThresholdPt: run?.kerningThresholdPt,
			}),
			...(leader ? { leader } : {}),
			...(run?.script ? { script: run.script } : {}),
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

export function alignFragments(
	fragments: LayoutFragment[],
	width: number,
	isLastLine: boolean,
	effectiveAlign: 'left' | 'right' | 'center' | 'justify',
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
