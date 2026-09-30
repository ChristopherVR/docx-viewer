import { describe, expect, it } from 'vitest';
import type { LayoutParagraph } from './input.js';
import type { TextMeasurer } from './measure.js';
import { fontOf, runTextWidth } from './paragraph-tokens.js';
import { layoutParagraph } from './paragraph-layout.js';

const measurer: TextMeasurer = {
	widthOf: (text) => [...text].length * 10,
	lineHeightOf: () => 20,
	ascentOf: () => 15,
};
const paragraph = (runs: LayoutParagraph['runs']): LayoutParagraph => ({
	kind: 'paragraph',
	id: 'advanced',
	runs,
});

describe('advanced font line layout', () => {
	it('adds spacing after scale, matching Word COM measurements', () => {
		expect(
			runTextWidth(
				'MMMM',
				{ text: 'MMMM', textScalePercent: 200, characterSpacingPx: 2 },
				measurer,
			),
		).toBe(88);
		expect(
			runTextWidth(
				'MMMM',
				{ text: 'MMMM', textScalePercent: 50, characterSpacingPx: -1 },
				measurer,
			),
		).toBe(16);
		expect(runTextWidth('MMMM', { text: 'MMMM', textScalePercent: 0 }, measurer)).toBe(0);
	});

	it('counts spacing by graphemes rather than splitting combining marks or emoji', () => {
		const constant: TextMeasurer = { ...measurer, widthOf: () => 10 };
		expect(runTextWidth('e\u0301👩‍💻', { text: '', characterSpacingPx: 2 }, constant)).toBe(14);
	});

	it('uses scaled/expanded widths for line wrapping and positions', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aaaa bbbb', textScalePercent: 200, characterSpacingPx: 2 }]),
			100,
			measurer,
			() => {},
		);
		expect(
			result.lines.map((line) => line.fragments.map((fragment) => fragment.text).join('')),
		).toEqual(['aaaa', 'bbbb']);
		expect(result.lines[0]!.fragments[0]).toMatchObject({
			widthPx: 88,
			textScalePercent: 200,
			characterSpacingPx: 2,
		});
	});

	it('uses scaled decimal prefixes when placing decimal tabs', () => {
		const p = paragraph([{ text: '\t12.34', textScalePercent: 200 }]);
		p.tabStops = [{ posPx: 100, align: 'decimal' }];
		const line = layoutParagraph(p, 300, measurer, () => {}).lines[0]!;
		expect(line.fragments[0]!.widthPx).toBe(60);
		expect(line.fragments[1]!.xPx).toBe(60);
	});

	it('uses the kerning threshold and explicit off without changing font size', () => {
		expect(fontOf({ text: 'AV', fontSizePt: 12, kerningThresholdPt: 12 })).toMatchObject({
			kerning: 'normal',
			sizePx: 16,
		});
		expect(fontOf({ text: 'AV', fontSizePt: 11, kerningThresholdPt: 12 }).kerning).toBe('none');
		expect(fontOf({ text: 'AV', fontSizePt: 24, kerningThresholdPt: 0 }).kerning).toBe('none');
		expect(fontOf({ text: 'AV', fontSizePt: 12 }).kerning).toBeUndefined();
	});

	it('raises and lowers full-size fragments around a shared baseline', () => {
		const line = layoutParagraph(
			paragraph([{ text: 'a', positionPx: 4 }, { text: 'b' }, { text: 'c', positionPx: -4 }]),
			200,
			measurer,
			() => {},
		).lines[0]!;
		expect(line.heightPx).toBe(28);
		expect(line.fragments.map((fragment) => fragment.topPx)).toEqual([0, 4, 8]);
		expect(line.fragments.map((fragment) => fragment.boxHeightPx)).toEqual([20, 20, 20]);
	});
});
