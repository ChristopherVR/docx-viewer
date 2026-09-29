import { describe, expect, expectTypeOf, it } from 'vitest';
import {
	eighthPoints,
	eighthPointsFromPoints,
	emu,
	emuFromPixels,
	halfPoints,
	halfPointsFromPoints,
	roundEighthPoints,
	roundHalfPoints,
	roundSignedTwips,
	roundTwips,
	signedTwips,
	signedTwipsFromPixels,
	signedTwipsFromPoints,
	twips,
	twipsFromInches,
	twipsFromPixels,
	twipsFromPoints,
	twipsToInches,
	twipsToPixels,
	twipsToPoints,
	type EighthPoints,
	type Emu,
	type HalfPoints,
	type Paragraph,
	type SectionProperties,
	type SignedTwips,
	type Table,
	type TableBorderSide,
	type Twips,
} from './index.js';

describe('strict constructors', () => {
	it('twips rejects negatives and fractions; signedTwips allows negatives only', () => {
		expect(twips(0)).toBe(0);
		expect(() => twips(-1)).toThrow(RangeError);
		expect(() => twips(1.5)).toThrow(RangeError);
		expect(signedTwips(-720)).toBe(-720);
		expect(() => signedTwips(0.5)).toThrow(RangeError);
		expect(() => signedTwips(Number.POSITIVE_INFINITY)).toThrow(RangeError);
		expect(() => eighthPoints(0.5)).toThrow(RangeError);
		expect(() => emu(Number.NaN)).toThrow(RangeError);
	});
});

describe('rounding constructors', () => {
	it('round to whole units, halves away from zero, without -0', () => {
		expect(roundTwips(10.4)).toBe(10);
		expect(roundTwips(10.5)).toBe(11);
		expect(roundSignedTwips(-10.5)).toBe(-11);
		expect(Object.is(roundSignedTwips(-0.2), 0)).toBe(true);
		expect(Object.is(roundTwips(-0.2), 0)).toBe(true);
		expect(roundHalfPoints(21.5)).toBe(22);
		expect(roundEighthPoints(3.6)).toBe(4);
	});
	it('roundTwips clamps negatives to zero while the signed variant keeps them', () => {
		expect(roundTwips(-40)).toBe(0);
		expect(roundSignedTwips(-40.4)).toBe(-40);
		expect(roundEighthPoints(-2)).toBe(0);
	});
	it('reject non-finite input instead of writing NaN into the model', () => {
		for (const bad of [Number.NaN, Infinity, -Infinity]) {
			expect(() => roundTwips(bad)).toThrow(RangeError);
			expect(() => roundSignedTwips(bad)).toThrow(RangeError);
			expect(() => roundHalfPoints(bad)).toThrow(RangeError);
		}
	});
});

describe('unit conversions', () => {
	it('convert from points, pixels and inches with rounding', () => {
		expect(twipsFromPoints(36)).toBe(720);
		expect(twipsFromPoints(0.26)).toBe(5);
		expect(signedTwipsFromPoints(-36)).toBe(-720);
		expect(twipsFromPixels(96)).toBe(1440);
		expect(twipsFromPixels(10.03)).toBe(150);
		expect(signedTwipsFromPixels(-7.5)).toBe(-113);
		expect(twipsFromInches(1.5)).toBe(2160);
		expect(halfPointsFromPoints(10.5)).toBe(21);
		expect(halfPointsFromPoints(10.26)).toBe(21);
		expect(eighthPointsFromPoints(0.5)).toBe(4);
		expect(emuFromPixels(96)).toBe(914400);
	});
	it('convert back to plain numbers', () => {
		expect(twipsToPoints(twips(720))).toBe(36);
		expect(twipsToPixels(signedTwips(-720))).toBe(-48);
		expect(twipsToInches(twips(2160))).toBe(1.5);
	});
	it('round-trips whole pixels through twips', () => {
		for (const px of [0, 1, 37, 816, 1056]) expect(twipsToPixels(twipsFromPixels(px))).toBe(px);
	});
});

describe('compile-time unit safety', () => {
	it('brands are distinct types that stay numbers at runtime', () => {
		expectTypeOf<Twips>().toMatchTypeOf<number>();
		expectTypeOf<Twips>().toMatchTypeOf<SignedTwips>();
		expectTypeOf<SignedTwips>().not.toMatchTypeOf<Twips>();
		expectTypeOf<HalfPoints>().not.toMatchTypeOf<Twips>();
		expectTypeOf<Twips>().not.toMatchTypeOf<HalfPoints>();
		expectTypeOf<EighthPoints>().not.toMatchTypeOf<HalfPoints>();
		expectTypeOf<Emu>().not.toMatchTypeOf<Twips>();
		expectTypeOf<number>().not.toMatchTypeOf<Twips>();
	});
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
