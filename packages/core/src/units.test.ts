import { describe, expect, expectTypeOf, it } from 'vitest';
import {
	eighthPoints,
	halfPoints,
	roundTwips,
	signedTwips,
	twips,
	type EighthPoints,
	type Paragraph,
	type SectionProperties,
	type SignedTwips,
	type Table,
	type TableBorderSide,
	type Twips,
} from './index.js';

// The unit types themselves are tested in ooxml-core (`@christophervr/ooxml-units`); these checks
// cover that the Word model's fields carry the brands.
describe('Word model unit safety', () => {
	it('model fields carry the brands', () => {
		expectTypeOf<Paragraph['spacingAfterTwips']>().toEqualTypeOf<Twips | undefined>();
		expectTypeOf<Paragraph['indentLeftTwips']>().toEqualTypeOf<SignedTwips | undefined>();
		expectTypeOf<Table['grid']>().toEqualTypeOf<Twips[] | undefined>();
		expectTypeOf<SectionProperties['pageWidthTwips']>().toEqualTypeOf<Twips>();
		expectTypeOf<SectionProperties['marginTopTwips']>().toEqualTypeOf<SignedTwips>();
		expectTypeOf<TableBorderSide['sizeEighthPoints']>().toEqualTypeOf<EighthPoints | undefined>();
	});
	it('rejects plain numbers and other units in model fields', () => {
		const paragraph: Paragraph = { type: 'paragraph', id: 'p', runs: [] };
		// @ts-expect-error a plain number is not a Twips
		paragraph.spacingAfterTwips = 240;
		// @ts-expect-error half-points are not twips
		paragraph.spacingAfterTwips = halfPoints(24);
		// @ts-expect-error eighth-points are not twips
		paragraph.spacingBeforeTwips = eighthPoints(4);
		// @ts-expect-error a signed measure cannot go into an unsigned field
		paragraph.hangingTwips = signedTwips(-360);
		// an unsigned measure is a valid signed one
		paragraph.indentLeftTwips = twips(720);
		paragraph.indentLeftTwips = signedTwips(-720);
		// @ts-expect-error a plain number is not an EighthPoints border size
		const side: TableBorderSide = { sizeEighthPoints: 4 };
		// @ts-expect-error twips are not eighth-points
		const other: TableBorderSide = { sizeEighthPoints: twips(4) };
		// @ts-expect-error a plain-number array is not a Twips[] grid
		const table: Table = { type: 'table', id: 't', rows: [], grid: [1500, 1500] };
		expect([side, other, table]).toHaveLength(3);
		expect(paragraph.indentLeftTwips).toBe(-720);
	});
	it('computed values must go through a rounding constructor', () => {
		const paragraph: Paragraph = { type: 'paragraph', id: 'p', runs: [] };
		const dragged = 100.4;
		// @ts-expect-error an unrounded computed number cannot reach the model
		paragraph.firstLineTwips = dragged;
		paragraph.firstLineTwips = roundTwips(dragged);
		expect(paragraph.firstLineTwips).toBe(100);
	});
});
