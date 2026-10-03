import type { RunFormatting } from 'docx-core';
import { cssFontStack, tokenizeRun, type TextMeasurer } from 'ooxml-core/docx/layout';
import { createCanvasMeasurer } from './canvas-measurer';

const graphemes = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
export interface ScaledSegment {
	from: number;
	to: number;
	css: string;
}

/** Reserve scaled advances while leaving source text available to the browser's caret/selection. */
export function scaledSegments(
	text: string,
	formatting: RunFormatting,
	family: string | undefined,
	measurer: TextMeasurer,
): ScaledSegment[] {
	const factor = (formatting.textScalePercent ?? 100) / 100;
	if (factor === 1) return [];
	const size = formatting.fontSize ?? 11;
	const spacing = (formatting.characterSpacingTwips ?? 0) / 15;
	const font = {
		family: family ?? 'Calibri',
		sizePx:
			((size * 4) / 3) *
			(formatting.verticalAlign === 'superscript' || formatting.verticalAlign === 'subscript'
				? 0.65
				: 1),
		...(formatting.bold ? { bold: true } : {}),
		...(formatting.italic ? { italic: true } : {}),
		...(formatting.smallCaps ? { smallCaps: true } : {}),
		...(formatting.ligatures ? { ligatures: formatting.ligatures } : {}),
		...(formatting.kerningHalfPoints === undefined
			? {}
			: {
					kerning:
						formatting.kerningHalfPoints > 0 && size >= formatting.kerningHalfPoints / 2
							? ('normal' as const)
							: ('none' as const),
				}),
	};
	return tokenizeRun(text, 0).flatMap((token) => {
		if (token.kind !== 'word' && token.kind !== 'space') return [];
		const shapedText = formatting.caps ? token.text.toUpperCase() : token.text;
		const glyphWidth = measurer.widthOf(shapedText, font);
		const count = [...graphemes.segment(shapedText)].length;
		const nativeSpacing = factor === 0 ? 0 : spacing / factor;
		const nativeWidth = Math.max(0, glyphWidth + count * nativeSpacing);
		const advance = Math.max(0, glyphWidth * factor + count * spacing);
		if (token.kind === 'space')
			return [
				{
					from: token.sourceStart,
					to: token.sourceStart + token.text.length,
					css: `display:inline;white-space:pre-wrap;font-family:${cssFontStack(font.family)};font-size:${font.sizePx}px;letter-spacing:${advance - glyphWidth}px`,
				},
			];
		return [
			{
				from: token.sourceStart,
				to: token.sourceStart + token.text.length,
				css: `display:inline-block;white-space:pre;font-family:${cssFontStack(font.family)};font-size:${font.sizePx}px;transform:scaleX(${factor});transform-origin:left center;letter-spacing:${nativeSpacing}px;margin-right:${advance - nativeWidth}px`,
			},
		];
	});
}

/** Each editor owns its cache; replacing it when fonts load avoids measuring fallback faces forever. */
export function scaleMeasurer() {
	return createCanvasMeasurer();
}
