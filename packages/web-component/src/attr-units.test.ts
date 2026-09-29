import { describe, expect, it } from 'vitest';
import { paragraphTwipsFromAttrs, signedTwipsAttr, twipsAttr } from './attr-units';

describe('editor attribute to model unit boundary', () => {
	it('rounds fractional values so they never reach the model as non-integers', () => {
		expect(twipsAttr(360.4)).toBe(360);
		expect(twipsAttr(360.5)).toBe(361);
		expect(signedTwipsAttr(-720.6)).toBe(-721);
	});

	it('drops values that are not valid lengths instead of writing them', () => {
		for (const bad of [null, undefined, '240', Number.NaN, Infinity]) {
			expect(twipsAttr(bad)).toBeUndefined();
			expect(signedTwipsAttr(bad)).toBeUndefined();
		}
		expect(twipsAttr(-1)).toBeUndefined();
		expect(signedTwipsAttr(-1)).toBe(-1);
	});

	it('maps paragraph attributes with the schema signedness and omits absent ones', () => {
		expect(
			paragraphTwipsFromAttrs({
				spacingBeforeTwips: 120.2,
				spacingAfterTwips: null,
				lineSpacingTwips: 276,
				indentLeftTwips: -360,
				indentStartTwips: null,
				firstLineTwips: -5,
				hangingTwips: 360,
			}),
		).toEqual({
			spacingBeforeTwips: 120,
			lineSpacingTwips: 276,
			indentLeftTwips: -360,
			hangingTwips: 360,
		});
	});
});
