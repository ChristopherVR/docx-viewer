// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { RunFormatting } from './run-style-model.js';
import type { ThemeColorToken } from './theme-model.js';
import { first, getW, type XmlElement } from './xml.js';
import { isWordUnderlineStyle } from './underline.js';

const on = (element: XmlElement | undefined): boolean => {
	if (!element) return false;
	const value = getW(element, 'val')?.toLowerCase();
	return !['0', 'false', 'off', 'no', 'none'].includes(value ?? '');
};
const points = (halfPoints: string | undefined): number | undefined =>
	halfPoints === undefined ? undefined : Number(halfPoints) / 2;
const twips = (value: string | undefined): number | undefined => {
	if (value === undefined || !/^-?\d+$/.test(value)) return undefined;
	const parsed = Number(value);
	return Number.isSafeInteger(parsed) ? parsed : undefined;
};

/** Shared `w:rPr` -> `RunFormatting` parsing, used for direct runs, docDefaults and style catalogs. */
export function parseRunProperties(props: XmlElement | undefined): RunFormatting {
	const result: RunFormatting = {};
	if (!props) return result;
	// Toggles are kept when explicitly off too (`w:val="0"`): that cancels a style's value.
	const toggle = (local: string): boolean | undefined => {
		const element = first(props, local);
		return element ? on(element) : undefined;
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
	if (underline && !on(underline)) result.underline = false;
	if (underline && on(underline)) {
		result.underline = true;
		const value = getW(underline, 'val')?.toLowerCase();
		if (value && isWordUnderlineStyle(value) && value !== 'single' && value !== 'none')
			result.underlineStyle = value;
		const color = getW(underline, 'color');
		if (color && /^[0-9a-f]{6}$/i.test(color)) result.underlineColor = `#${color}`;
	}
	const highlight = getW(first(props, 'highlight'), 'val');
	if (highlight) result.highlight = highlight;
	const verticalAlign = getW(first(props, 'vertAlign'), 'val');
	if (verticalAlign === 'superscript' || verticalAlign === 'subscript')
		result.verticalAlign = verticalAlign;
	const size = points(getW(first(props, 'sz'), 'val'));
	if (size !== undefined) result.fontSize = size;
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
	const hex = getW(color, 'val');
	if (hex && /^[0-9a-f]{6}$/i.test(hex)) result.color = `#${hex}`;
	const themeColor = getW(color, 'themeColor');
	if (themeColor) {
		result.colorTheme = { token: themeColor as ThemeColorToken };
		const tint = getW(color, 'themeTint');
		const shade = getW(color, 'themeShade');
		if (tint && /^[0-9a-fA-F]{2}$/.test(tint))
			result.colorTheme.tint = Number.parseInt(tint, 16) / 255;
		if (shade && /^[0-9a-fA-F]{2}$/.test(shade))
			result.colorTheme.shade = Number.parseInt(shade, 16) / 255;
	}
	const spacing = twips(getW(first(props, 'spacing'), 'val'));
	if (spacing !== undefined) result.characterSpacingTwips = spacing;
	const shd = first(props, 'shd');
	const fill = getW(shd, 'fill');
	if (fill && /^[0-9a-f]{6}$/i.test(fill)) result.shadingFill = `#${fill}`;
	const shdTheme = getW(shd, 'themeFill');
	if (shdTheme) {
		result.shadingThemeFill = { token: shdTheme as ThemeColorToken };
		const tint = getW(shd, 'themeFillTint');
		const shade = getW(shd, 'themeFillShade');
		if (tint && /^[0-9a-fA-F]{2}$/.test(tint))
			result.shadingThemeFill.tint = Number.parseInt(tint, 16) / 255;
		if (shade && /^[0-9a-fA-F]{2}$/.test(shade))
			result.shadingThemeFill.shade = Number.parseInt(shade, 16) / 255;
	}
	return result;
}
