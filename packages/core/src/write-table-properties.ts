import type { Table } from './model.js';
import type { TableBorders, TableRowProperties } from './table-model.js';
import { buildRowProperties } from './table-cell-write.js';
import { buildBorders, buildMargins } from './table-defaults.js';
import { orderChildren } from './element-order.js';
import { children, first, makeW, type XmlDocument, type XmlElement } from './xml.js';

const ROW_ORDER = [
	'cnfStyle',
	'divId',
	'gridBefore',
	'gridAfter',
	'wBefore',
	'wAfter',
	'cantSplit',
	'trHeight',
	'tblHeader',
	'tblCellSpacing',
	'jc',
	'hidden',
	'ins',
	'del',
	'trPrChange',
];
const TABLE_ORDER = [
	'tblStyle',
	'tblpPr',
	'tblOverlap',
	'bidiVisual',
	'tblStyleRowBandSize',
	'tblStyleColBandSize',
	'tblW',
	'jc',
	'tblCellSpacing',
	'tblInd',
	'tblBorders',
	'shd',
	'tblLayout',
	'tblCellMar',
	'tblLook',
	'tblCaption',
	'tblDescription',
	'tblPrChange',
];

/** Changes only modeled row fields, preserving unmodeled trPr children and attributes. */
export function patchRowProperties(
	doc: XmlDocument,
	tr: XmlElement,
	next: TableRowProperties | undefined,
	base: TableRowProperties | undefined,
): void {
	if (JSON.stringify(next ?? {}) === JSON.stringify(base ?? {})) return;
	let props = first(tr, 'trPr');
	if (!props) {
		props = makeW(doc, 'trPr');
		tr.insertBefore(props, tr.firstChild);
	}
	const built = buildRowProperties(doc, next, 0);
	for (const [name, changed] of [
		['cantSplit', Boolean(next?.cantSplit) !== Boolean(base?.cantSplit)],
		['tblHeader', Boolean(next?.header) !== Boolean(base?.header)],
		['trHeight', next?.heightTwips !== base?.heightTwips || next?.heightRule !== base?.heightRule],
	] as const) {
		if (!changed) continue;
		for (const old of children(props, name)) props.removeChild(old);
		const replacement = built && first(built, name);
		if (replacement) props.appendChild(replacement);
	}
	orderChildren(props, ROW_ORDER);
}

/** Updates table defaults without replacing borders, style, look or extension XML. */
export function patchTableMargins(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base?: Table,
): void {
	if (JSON.stringify(table.cellMargins ?? {}) === JSON.stringify(base?.cellMargins ?? {})) return;
	let props = first(node, 'tblPr');
	if (!props) {
		props = makeW(doc, 'tblPr');
		node.insertBefore(props, node.firstChild);
	}
	let margins = first(props, 'tblCellMar');
	if (!margins) {
		margins = makeW(doc, 'tblCellMar');
		props.appendChild(margins);
	}
	const built = buildMargins(doc, table.cellMargins ?? {}, 'tblCellMar');
	for (const side of ['top', 'left', 'bottom', 'right'] as const) {
		if (table.cellMargins?.[side] === base?.cellMargins?.[side]) continue;
		for (const old of children(margins, side)) margins.removeChild(old);
		const replacement = first(built, side);
		if (replacement) margins.appendChild(replacement);
	}
	orderChildren(margins, ['top', 'left', 'start', 'bottom', 'right', 'end']);
	orderChildren(props, TABLE_ORDER);
}

const CELL_ORDER = [
	'cnfStyle',
	'tcW',
	'gridSpan',
	'hMerge',
	'vMerge',
	'tcBorders',
	'shd',
	'noWrap',
	'tcMar',
	'textDirection',
	'tcFitText',
	'vAlign',
	'hideMark',
	'headers',
	'cellIns',
	'cellDel',
	'cellMerge',
	'tcPrChange',
];
const CELL_BORDER_ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV'];
const CELL_SIDES = ['top', 'left', 'bottom', 'right'] as const;

/**
 * Changes only the cell border sides that differ from `base`, keeping every other `tcPr` child
 * (width, shading, margins, extensions) and untouched sides exactly as they were.
 */
export function patchCellBorders(
	doc: XmlDocument,
	tc: XmlElement,
	next: TableBorders | undefined,
	base: TableBorders | undefined,
): void {
	const changed = CELL_SIDES.filter(
		(side) => JSON.stringify(next?.[side]) !== JSON.stringify(base?.[side]),
	);
	if (!changed.length) return;
	let props = first(tc, 'tcPr');
	if (!props) {
		props = makeW(doc, 'tcPr');
		tc.insertBefore(props, tc.firstChild);
	}
	let borders = first(props, 'tcBorders');
	if (!borders) {
		borders = makeW(doc, 'tcBorders');
		props.appendChild(borders);
	}
	const built = buildBorders(doc, next ?? {}, 'tcBorders');
	for (const side of changed) {
		for (const old of children(borders, side)) borders.removeChild(old);
		const replacement = first(built, side);
		if (replacement) borders.appendChild(replacement);
	}
	if (!Array.from(borders.childNodes).some((node) => node.nodeType === 1))
		props.removeChild(borders);
	else orderChildren(borders, CELL_BORDER_ORDER);
	orderChildren(props, CELL_ORDER);
	if (!Array.from(props.childNodes).some((node) => node.nodeType === 1)) tc.removeChild(props);
}
