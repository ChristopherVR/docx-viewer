// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { RunFormatting } from './run-style-model.js';
import { isStHighlightColor, isStThemeColor } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import {
	onOffElement,
	parseHalfPoints,
	parseRgbColor,
	parseSignedTwips,
	parseTintShade,
} from './simple-types.js';
import { first, getW, type XmlElement } from './xml.js';
import { isWordUnderlineStyle } from './underline.js';

/** Shared `w:rPr` -> `RunFormatting` parsing, used for direct runs, docDefaults and style catalogs. */
export function parseRunProperties(props: XmlElement | undefined): RunFormatting {
	const result: RunFormatting = {};
	if (!props) return result;
	// Toggles are kept when explicitly off too (`w:val="0"`): that cancels a style's value.
	const toggle = (local: string): boolean | undefined => {
		return onOffElement(first(props, local));
	};
	for (const [local, key] of [
		['b', 'bold'],
		['i', 'italic'],
		['caps', 'caps'],
		['smallCaps', 'smallCaps'],
		['vanish', 'vanish'],
		['strike', 'strike'],
		['dstrike', 'doubleStrike'],
	] as const) {
		const value = toggle(local);
		if (value !== undefined) result[key] = value;
	}
	const underline = first(props, 'u');
	// `w:u` carries an underline style (`ST_Underline`) in `w:val`, not an on/off value.
	const underlineValue = getW(underline, 'val')?.toLowerCase();
	if (
		underline &&
		(underlineValue === 'none' ||
			underlineValue === '0' ||
			underlineValue === 'false' ||
			underlineValue === 'off')
	)
		result.underline = false;
	else if (underline) {
		result.underline = true;
		if (underlineValue && isWordUnderlineStyle(underlineValue) && underlineValue !== 'single')
			result.underlineStyle = underlineValue;
		const color = parseRgbColor(getW(underline, 'color'));
		if (color) result.underlineColor = color;
	}
	const highlight = enumValue(
		isStHighlightColor,
		getW(first(props, 'highlight'), 'val'),
		'w:highlight',
	);
	if (highlight) result.highlight = highlight;
	const verticalAlign = getW(first(props, 'vertAlign'), 'val');
	if (verticalAlign === 'superscript' || verticalAlign === 'subscript')
		result.verticalAlign = verticalAlign;
	const size = parseHalfPoints(getW(first(props, 'sz'), 'val'));
	if (size !== undefined) result.fontSize = size / 2;
	const fonts = first(props, 'rFonts');
	const family = getW(fonts, 'ascii') ?? getW(fonts, 'hAnsi');
	if (family) result.fontFamily = family;
	const fontTheme: RunFormatting['fontTheme'] = {};
	for (const [xmlKey, script] of [
		['asciiTheme', 'ascii'],
		['hAnsiTheme', 'hAnsi'],
		['eastAsiaTheme', 'eastAsia'],
		['cstheme', 'cs'],
	] as const) {
		const value = getW(fonts, xmlKey);
		if (
			value === 'majorHAnsi' ||
			value === 'majorEastAsia' ||
			value === 'majorBidi' ||
			value === 'majorAscii'
		)
			fontTheme[script] = 'major';
		else if (
			value === 'minorHAnsi' ||
			value === 'minorEastAsia' ||
			value === 'minorBidi' ||
			value === 'minorAscii'
		)
			fontTheme[script] = 'minor';
	}
	if (Object.keys(fontTheme).length) result.fontTheme = fontTheme;
	const color = first(props, 'color');
	const hex = parseRgbColor(getW(color, 'val'));
	if (hex) result.color = hex;
	const themeColor = enumValue(isStThemeColor, getW(color, 'themeColor'), 'w:themeColor');
	if (themeColor) {
		result.colorTheme = { token: themeColor };
		const tint = parseTintShade(getW(color, 'themeTint'));
		const shade = parseTintShade(getW(color, 'themeShade'));
		if (tint !== undefined) result.colorTheme.tint = tint;
		if (shade !== undefined) result.colorTheme.shade = shade;
	}
	const spacing = parseSignedTwips(getW(first(props, 'spacing'), 'val'));
	if (spacing !== undefined) result.characterSpacingTwips = spacing;
	const shd = first(props, 'shd');
	const fill = parseRgbColor(getW(shd, 'fill'));
	if (fill) result.shadingFill = fill;
	const shdTheme = enumValue(isStThemeColor, getW(shd, 'themeFill'), 'w:themeFill');
	if (shdTheme) {
		result.shadingThemeFill = { token: shdTheme };
		const tint = parseTintShade(getW(shd, 'themeFillTint'));
		const shade = parseTintShade(getW(shd, 'themeFillShade'));
		if (tint !== undefined) result.shadingThemeFill.tint = tint;
		if (shade !== undefined) result.shadingThemeFill.shade = shade;
	}
	return result;
}
