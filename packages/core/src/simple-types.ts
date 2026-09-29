// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Shared parsers for ECMA-376 simple types (shared-commonSimpleTypes.xsd and wml.xsd). Every
// parser returns undefined for values the schema does not allow, so callers never keep garbage.
import {
	eighthPoints,
	halfPoints,
	signedTwips,
	twips,
	type EighthPoints,
	type HalfPoints,
	type SignedTwips,
	type Twips,
} from './units.js';
import { getW, type XmlElement } from './xml.js';

/** `ST_OnOff`: exactly true|false|on|off|1|0. */
export function parseOnOff(value: string | null | undefined): boolean | undefined {
	switch (value) {
		case 'true':
		case 'on':
		case '1':
			return true;
		case 'false':
		case 'off':
		case '0':
			return false;
		default:
			return undefined;
	}
}

/**
 * `CT_OnOff` element: absent element -> undefined (inherit); present without `w:val` -> true;
 * an invalid `w:val` -> undefined rather than a guess.
 */
export function onOffElement(element: XmlElement | null | undefined): boolean | undefined {
	if (!element) return undefined;
	const value = getW(element, 'val');
	return value === undefined ? true : parseOnOff(value);
}

/** A `CT_OnOff`-typed attribute (such as `w:firstRow` on `w:tblLook`); absent or invalid -> undefined. */
export const onOffAttribute = (element: XmlElement, name: string): boolean | undefined =>
	parseOnOff(getW(element, name));

const SAFE_INTEGER = /^-?\d+$/;
/** `xsd:integer` that fits a safe JS integer. */
export function parseInteger(value: string | null | undefined): number | undefined {
	if (value === null || value === undefined || !SAFE_INTEGER.test(value)) return undefined;
	const parsed = Number(value);
	return Number.isSafeInteger(parsed) ? parsed : undefined;
}
/** `ST_DecimalNumber`-style non-negative integer. */
export function parseUnsignedInteger(value: string | null | undefined): number | undefined {
	const parsed = parseInteger(value);
	return parsed !== undefined && parsed >= 0 ? parsed : undefined;
}

const TWIPS_PER_UNIT = { in: 1440, cm: 1440 / 2.54, mm: 144 / 2.54, pt: 20, pc: 240, pi: 240 };
const UNIVERSAL_MEASURE = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/;

/** `ST_UniversalMeasure` (`1in`, `2.5cm`, `-3pt`...) converted to whole twips. */
export function universalMeasureToTwips(value: string): number | undefined {
	const match = UNIVERSAL_MEASURE.exec(value);
	if (!match) return undefined;
	const total = Number(match[1]) * TWIPS_PER_UNIT[match[2] as keyof typeof TWIPS_PER_UNIT];
	const rounded = Math.sign(total) * Math.round(Math.abs(total));
	return Number.isSafeInteger(rounded) ? rounded || 0 : undefined;
}

/** `ST_SignedTwipsMeasure`: an integer or a universal measure, in twips. */
export function parseSignedTwips(value: string | null | undefined): SignedTwips | undefined {
	if (value === null || value === undefined) return undefined;
	const text = value.trim();
	const parsed = SAFE_INTEGER.test(text) ? parseInteger(text) : universalMeasureToTwips(text);
	return parsed === undefined ? undefined : signedTwips(parsed);
}
/** `ST_TwipsMeasure`: like `parseSignedTwips` but negative values are rejected. */
export function parseTwips(value: string | null | undefined): Twips | undefined {
	const parsed = parseSignedTwips(value);
	return parsed !== undefined && parsed >= 0 ? twips(parsed) : undefined;
}

/** `ST_HpsMeasure`: non-negative half-points, or a positive universal measure (`12pt`). */
export function parseHalfPoints(value: string | null | undefined): HalfPoints | undefined {
	if (value === null || value === undefined) return undefined;
	const text = value.trim();
	if (/^\d+$/.test(text)) {
		const parsed = parseInteger(text);
		return parsed === undefined ? undefined : halfPoints(parsed);
	}
	const asTwips = UNIVERSAL_MEASURE.test(text) ? universalMeasureToTwips(text) : undefined;
	return asTwips !== undefined && asTwips >= 0 ? halfPoints(Math.round(asTwips / 10)) : undefined;
}
/** `ST_EighthPointMeasure`: non-negative integer eighths of a point. */
export function parseEighthPoints(value: string | null | undefined): EighthPoints | undefined {
	const parsed = parseUnsignedInteger(value?.trim());
	return parsed === undefined ? undefined : eighthPoints(parsed);
}

/** `ST_HexColor`: `auto` or six hex digits, returned as `#RRGGBB` (case preserved). */
export type HexColor = 'auto' | `#${string}`;
export function parseHexColor(value: string | null | undefined): HexColor | undefined {
	if (value === null || value === undefined) return undefined;
	if (value === 'auto') return 'auto';
	return /^[0-9a-fA-F]{6}$/.test(value) ? `#${value}` : undefined;
}
/** Like `parseHexColor` for the many attributes where `auto` carries no explicit color. */
export function parseRgbColor(value: string | null | undefined): `#${string}` | undefined {
	const parsed = parseHexColor(value);
	return parsed === 'auto' ? undefined : parsed;
}
/** `ST_UcharHexNumber` tint/shade (two hex digits) as a 0..1 fraction. */
export function parseTintShade(value: string | null | undefined): number | undefined {
	return value && /^[0-9a-fA-F]{2}$/.test(value) ? Number.parseInt(value, 16) / 255 : undefined;
}
