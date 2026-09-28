// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { ParagraphBorders, TableBorderSide, TableBorders } from './table-model.js';
import type { ThemeColorToken } from './theme-model.js';
import { isStBorder, isStThemeColor } from './generated/wml-simple-types.js';
import { enumValue } from './parse-diagnostics.js';
import {
	parseEighthPoints,
	parseRgbColor,
	parseTintShade,
	parseUnsignedInteger,
} from './simple-types.js';
import { first, getW, type XmlElement } from './xml.js';

const SIDES = [
	'top',
	'bottom',
	'left',
	'right',
	'insideH',
	'insideV',
] as const satisfies readonly (keyof TableBorders)[];

export function borderSide(element: XmlElement | undefined): TableBorderSide | undefined {
	if (!element) return undefined;
	const style = enumValue(isStBorder, getW(element, 'val'), 'w:val (border style)');
	const size = parseEighthPoints(getW(element, 'sz'));
	const color = parseRgbColor(getW(element, 'color'));
	const themeColor = enumValue(isStThemeColor, getW(element, 'themeColor'), 'w:themeColor');
	if (!style && size === undefined && !color && !themeColor) return undefined;
	const side: TableBorderSide = {};
	if (style) side.style = style;
	if (size !== undefined) side.sizeEighthPoints = size;
	if (color) side.color = color;
	if (themeColor) side.themeColor = themeColor;
	return side;
}

/** Parses a `w:tblBorders` or `w:tcBorders` element into the shared `TableBorders` shape. */
export function parseTableBorders(
	bordersElement: XmlElement | undefined,
): TableBorders | undefined {
	if (!bordersElement) return undefined;
	const result: TableBorders = {};
	// Word 2010+ writes `w:start`/`w:end` for the left/right edges of left-to-right tables.
	const alias: Partial<Record<(typeof SIDES)[number], string>> = { left: 'start', right: 'end' };
	for (const side of SIDES) {
		const aliasName = alias[side];
		const value = borderSide(
			first(bordersElement, side) ?? (aliasName ? first(bordersElement, aliasName) : undefined),
		);
		if (value) result[side] = value;
	}
	return Object.keys(result).length ? result : undefined;
}

export function parseShadingFill(shd: XmlElement | undefined): string | undefined {
	return parseRgbColor(getW(shd, 'fill'));
}
export function parseShadingThemeFill(shd: XmlElement | undefined) {
	const theme = enumValue(isStThemeColor, getW(shd, 'themeFill'), 'w:themeFill');
	if (!theme) return undefined;
	const ref: { token: ThemeColorToken; tint?: number; shade?: number } = { token: theme };
	const tint = parseTintShade(getW(shd, 'themeFillTint'));
	const shade = parseTintShade(getW(shd, 'themeFillShade'));
	if (tint !== undefined) ref.tint = tint;
	if (shade !== undefined) ref.shade = shade;
	return ref;
}

/** Parses `w:pBdr`: box sides, the line `between` paragraphs of a group, and each side's `w:space` (points). */
export function parseParagraphBorders(pBdr: XmlElement | undefined): ParagraphBorders | undefined {
	if (!pBdr) return undefined;
	const result: ParagraphBorders = {};
	for (const side of ['top', 'bottom', 'left', 'right', 'between'] as const) {
		const element =
			first(pBdr, side) ??
			(side === 'left' ? first(pBdr, 'start') : side === 'right' ? first(pBdr, 'end') : undefined);
		const value = borderSide(element);
		if (!value) continue;
		const space = parseUnsignedInteger(getW(element, 'space'));
		if (space !== undefined) value.spacePoints = space;
		result[side] = value;
	}
	return Object.keys(result).length ? result : undefined;
}
