// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Serializes cell (`w:tcPr`) and row (`w:trPr`) properties for tables the editor creates, so
// nothing the model records about a new table is silently dropped.
import type { TableCell } from './model.js';
import type { TableRowProperties } from './table-model.js';
import { fractionToThemeByte } from './theme-color.js';
import { buildBorders, buildMargins } from './table-defaults.js';
import { makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

const setW = (element: XmlElement, name: string, value: string): void =>
	element.setAttributeNS(WORD_NS, `w:${name}`, value);

function withVal(doc: XmlDocument, name: string, value: string): XmlElement {
	const element = makeW(doc, name);
	setW(element, 'val', value);
	return element;
}

/**
 * Builds `w:tcPr` in CT_TcPr order (wml.xsd): tcW, gridSpan, vMerge, tcBorders, shd, noWrap,
 * tcMar, textDirection, tcFitText, vAlign, hideMark. The model records none of noWrap,
 * textDirection, tcFitText or hideMark.
 */
export function buildCellProperties(doc: XmlDocument, cell: TableCell): XmlElement | undefined {
	const tcPr = makeW(doc, 'tcPr');
	if (cell.widthTwips !== undefined) {
		const width = makeW(doc, 'tcW');
		setW(width, 'w', String(cell.widthTwips));
		setW(width, 'type', 'dxa');
		tcPr.appendChild(width);
	}
	if (cell.gridSpan !== undefined && cell.gridSpan > 1)
		tcPr.appendChild(withVal(doc, 'gridSpan', String(cell.gridSpan)));
	if (cell.verticalMerge) tcPr.appendChild(withVal(doc, 'vMerge', cell.verticalMerge));
	if (cell.borders) {
		const borders = buildBorders(doc, cell.borders, 'tcBorders');
		if (borders.childNodes.length) tcPr.appendChild(borders);
	}
	if (cell.shadingFill || cell.shadingThemeFill) {
		const shd = makeW(doc, 'shd');
		setW(shd, 'val', 'clear');
		setW(shd, 'color', 'auto');
		setW(shd, 'fill', cell.shadingFill ? cell.shadingFill.replace(/^#/, '') : 'auto');
		const theme = cell.shadingThemeFill;
		if (theme) {
			setW(shd, 'themeFill', theme.token);
			if (theme.tint !== undefined) setW(shd, 'themeFillTint', fractionToThemeByte(theme.tint));
			if (theme.shade !== undefined) setW(shd, 'themeFillShade', fractionToThemeByte(theme.shade));
		}
		tcPr.appendChild(shd);
	}
	const margins = cell.margins && buildMargins(doc, cell.margins, 'tcMar');
	if (margins) tcPr.appendChild(margins);
	if (cell.verticalAlign) tcPr.appendChild(withVal(doc, 'vAlign', cell.verticalAlign));
	return tcPr.childNodes.length ? tcPr : undefined;
}

/** Builds `w:trPr`; `gridAfter` records grid columns a short row leaves unused. */
export function buildRowProperties(
	doc: XmlDocument,
	row: TableRowProperties | undefined,
	gridAfter: number,
): XmlElement | undefined {
	const trPr = makeW(doc, 'trPr');
	if (gridAfter > 0) trPr.appendChild(withVal(doc, 'gridAfter', String(gridAfter)));
	if (row?.cantSplit) trPr.appendChild(makeW(doc, 'cantSplit'));
	if (row?.heightTwips !== undefined) {
		const height = makeW(doc, 'trHeight');
		setW(height, 'val', String(row.heightTwips));
		setW(height, 'hRule', row.heightRule ?? 'atLeast');
		trPr.appendChild(height);
	}
	if (row?.header) trPr.appendChild(makeW(doc, 'tblHeader'));
	return trPr.childNodes.length ? trPr : undefined;
}
