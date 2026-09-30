import type { Table } from './model.js';
import type { TableRowProperties } from './table-model.js';
import { buildRowProperties } from './table-cell-write.js';
import { buildMargins } from './table-defaults.js';
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
