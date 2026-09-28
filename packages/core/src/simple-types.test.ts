import { describe, expect, it } from 'vitest';
import {
	onOffElement,
	parseEighthPoints,
	parseHalfPoints,
	parseHexColor,
	parseOnOff,
	parseRgbColor,
	parseSignedTwips,
	parseTwips,
	universalMeasureToTwips,
} from './simple-types.js';
import { emu, halfPoints, twips } from './units.js';
import { parseXml } from './xml.js';

const element = (attributes: string) =>
	parseXml(
		`<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ${attributes}/>`,
	).documentElement;

describe('parseOnOff (ST_OnOff)', () => {
	it('accepts exactly true, false, on, off, 1 and 0', () => {
		for (const value of ['true', 'on', '1']) expect(parseOnOff(value)).toBe(true);
		for (const value of ['false', 'off', '0']) expect(parseOnOff(value)).toBe(false);
	});
	it('rejects garbage instead of treating it as true', () => {
		for (const value of ['yes', 'no', 'TRUE', '', '2', 'none', ' true'])
			expect(parseOnOff(value)).toBeUndefined();
		expect(parseOnOff(undefined)).toBeUndefined();
		expect(parseOnOff(null)).toBeUndefined();
	});
	it('treats a present element without w:val as true and garbage as absent', () => {
		expect(onOffElement(undefined)).toBeUndefined();
		expect(onOffElement(element(''))).toBe(true);
		expect(onOffElement(element('w:val="off"'))).toBe(false);
		expect(onOffElement(element('w:val="banana"'))).toBeUndefined();
	});
});

describe('twips (ST_TwipsMeasure / ST_SignedTwipsMeasure)', () => {
	it('parses integers', () => {
		expect(parseTwips('1440')).toBe(1440);
		expect(parseSignedTwips('-720')).toBe(-720);
		expect(parseTwips('-720')).toBeUndefined();
	});
	it('converts universal measures', () => {
		expect(parseTwips('1in')).toBe(1440);
		expect(parseTwips('0.5in')).toBe(720);
		expect(parseTwips('2.54cm')).toBe(1440);
		expect(parseTwips('10mm')).toBe(567);
		expect(parseTwips('12pt')).toBe(240);
		expect(parseTwips('1pc')).toBe(240);
		expect(parseTwips('1pi')).toBe(240);
		expect(parseSignedTwips('-1in')).toBe(-1440);
		expect(parseTwips('-1in')).toBeUndefined();
		expect(Object.is(universalMeasureToTwips('-0in'), 0)).toBe(true);
	});
	it('rejects garbage', () => {
		for (const value of ['', 'abc', '1 in', '1ft', '1.5', '12px', '1e3', '--1', '50%', 'NaN'])
			expect(parseSignedTwips(value)).toBeUndefined();
		expect(parseTwips(undefined)).toBeUndefined();
		expect(parseTwips('99999999999999999999')).toBeUndefined();
	});
});

describe('font and border measures', () => {
	it('parses half-points and positive universal measures', () => {
		expect(parseHalfPoints('24')).toBe(24);
		expect(parseHalfPoints('12pt')).toBe(24);
		expect(parseHalfPoints('0.5in')).toBe(72);
		for (const value of ['-2', '', 'abc', '1.5', '-1pt', 'NaN'])
			expect(parseHalfPoints(value)).toBeUndefined();
	});
	it('parses eighth-points', () => {
		expect(parseEighthPoints('4')).toBe(4);
		expect(parseEighthPoints('-4')).toBeUndefined();
		expect(parseEighthPoints('1pt')).toBeUndefined();
	});
});

describe('parseHexColor (ST_HexColor)', () => {
	it('accepts auto and six hex digits', () => {
		expect(parseHexColor('auto')).toBe('auto');
		expect(parseHexColor('28665E')).toBe('#28665E');
		expect(parseHexColor('ff00aa')).toBe('#ff00aa');
		expect(parseRgbColor('auto')).toBeUndefined();
		expect(parseRgbColor('00FF00')).toBe('#00FF00');
	});
	it('rejects garbage', () => {
		for (const value of ['', '#FF0000', 'FFF', 'GGGGGG', 'FF00000', 'Auto'])
			expect(parseHexColor(value)).toBeUndefined();
	});
});

describe('branded units', () => {
	it('construct integers only and stay numbers at runtime', () => {
		expect(twips(20) + 1).toBe(21);
		expect(halfPoints(24)).toBe(24);
		expect(emu(914400)).toBe(914400);
		expect(() => twips(1.5)).toThrow(RangeError);
		expect(() => halfPoints(Number.NaN)).toThrow(RangeError);
	});
});
