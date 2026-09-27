// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TableBorderSide, TableBorders } from './table-model.js';
import type { ThemeColorToken } from './theme-model.js';
import { first, getW, type XmlElement } from './xml.js';

const SIDES = [
	'top',
	'bottom',
	'left',
	'right',
	'insideH',
	'insideV',
] as const satisfies readonly (keyof TableBorders)[];

function borderSide(element: XmlElement | undefined): TableBorderSide | undefined {
	if (!element) return undefined;
	const style = getW(element, 'val');
	const size = getW(element, 'sz');
	const color = getW(element, 'color');
	const themeColor = getW(element, 'themeColor');
	if (!style && !size && !color && !themeColor) return undefined;
	const side: TableBorderSide = {};
	if (style) side.style = style;
	if (size !== undefined && /^\d+$/.test(size)) side.sizeEighthPoints = Number(size);
	if (color && /^[0-9a-f]{6}$/i.test(color)) side.color = `#${color}`;
	if (themeColor) side.themeColor = themeColor as ThemeColorToken;
	return side;
}

/** Parses a `w:tblBorders` or `w:tcBorders` element into the shared `TableBorders` shape. */
export function parseTableBorders(
	bordersElement: XmlElement | undefined,
): TableBorders | undefined {
	if (!bordersElement) return undefined;
	const result: TableBorders = {};
	for (const side of SIDES) {
		const value = borderSide(first(bordersElement, side));
		if (value) result[side] = value;
	}
	return Object.keys(result).length ? result : undefined;
}

export function parseShadingFill(shd: XmlElement | undefined): string | undefined {
	const fill = getW(shd, 'fill');
	return fill && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill}` : undefined;
}
export function parseShadingThemeFill(shd: XmlElement | undefined) {
	const theme = getW(shd, 'themeFill');
	if (!theme) return undefined;
	const ref: { token: ThemeColorToken; tint?: number; shade?: number } = {
		token: theme as ThemeColorToken,
	};
	const tint = getW(shd, 'themeFillTint');
	const shade = getW(shd, 'themeFillShade');
	if (tint && /^[0-9a-fA-F]{2}$/.test(tint)) ref.tint = Number.parseInt(tint, 16) / 255;
	if (shade && /^[0-9a-fA-F]{2}$/.test(shade)) ref.shade = Number.parseInt(shade, 16) / 255;
	return ref;
}
