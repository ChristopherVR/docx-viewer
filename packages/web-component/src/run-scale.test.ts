// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createFakeMeasurer } from '@christophervr/ooxml-core/docx/layout';
import { signedTwips } from '@christophervr/docx-core';
import { scaledSegments } from './run-scale';

const measurer = createFakeMeasurer({ charWidthFactor: 0.5 });
const css = (text: string, scale: number, spacing = 0) => {
	const segment = scaledSegments(
		text,
		{ textScalePercent: scale, fontSize: 12, characterSpacingTwips: signedTwips(spacing) },
		'Courier New',
		measurer,
	)[0]!;
	const span = document.createElement('span');
	span.style.cssText = segment.css;
	return span.style;
};
describe('continuous text scaling', () => {
	it('reserves glyph scale plus spacing after scale, including zero scale', () => {
		expect(css('MMMM', 200, 30).marginRight).toBe('36px');
		expect(css('MMMM', 50, 30).marginRight).toBe('-24px');
		expect(css('MMMM', 0, 30).marginRight).toBe('-24px');
		expect(css('MMMM', 200, 30).letterSpacing).toBe('1px');
	});
	it('clamps severely condensed advances without creating positive layout space', () => {
		expect(css('MMMM', 50, -600).marginRight).toBe('0px');
	});
	it('keeps combining marks and emoji intact and retains UTF-16 source offsets', () => {
		const text = 'あ\u3099👩‍💻 text';
		const segments = scaledSegments(text, { textScalePercent: 125 }, undefined, measurer);
		expect(segments.map((segment) => text.slice(segment.from, segment.to))).toEqual([
			'あ\u3099',
			'👩‍💻',
			' ',
			'text',
		]);
		expect(segments[1]!.from).toBe(2);
		expect(segments[2]!.from).toBe(7);
	});
	it('leaves tabs and line breaks native and does not decorate neutral scale', () => {
		const text = 'one\ttwo\nthree';
		expect(
			scaledSegments(text, { textScalePercent: 200 }, undefined, measurer).map((segment) =>
				text.slice(segment.from, segment.to),
			),
		).toEqual(['one', 'two', 'three']);
		expect(scaledSegments(text, { textScalePercent: 100 }, undefined, measurer)).toEqual([]);
	});
});
