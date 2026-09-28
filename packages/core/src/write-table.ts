import { expectDefined } from './expect-defined.js';
import type { Paragraph, Table, TableCell } from './model.js';
import {
	children,
	first,
	isElement,
	makeW,
	named,
	type XmlDocument,
	type XmlElement,
} from './xml.js';

type ParagraphWriter = (
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base?: Paragraph,
) => XmlElement;
type SlotsWriter = (
	doc: XmlDocument,
	parent: XmlElement,
	old: XmlElement[],
	next: XmlElement[],
) => void;

/** Structural edits are restricted to plain rectangular tables with no hidden inline content. */
export function canEditTableStructure(table: XmlElement): boolean {
	const only = (parent: XmlElement, allowed: string[]) =>
		Array.from(parent.childNodes)
			.filter(isElement)
			.every((node) => allowed.some((name) => named(node, name)));
	if (!only(table, ['tblPr', 'tblGrid', 'tr'])) return false;
	const grid = first(table, 'tblGrid');
	if (grid && !only(grid, ['gridCol'])) return false;
	const rows = children(table, 'tr');
	if (!rows.length) return false;
	const [firstRow] = rows;
	if (!firstRow) return false;
	const width = children(firstRow, 'tc').length;
	if (!width) return false;
	if (grid && children(grid, 'gridCol').length !== width) return false;
	for (const row of rows) {
		if (!only(row, ['trPr', 'tc']) || children(row, 'tc').length !== width) return false;
		const rowProps = first(row, 'trPr');
		if (rowProps && ['gridBefore', 'gridAfter', 'trPrChange'].some((key) => first(rowProps, key)))
			return false;
		for (const cell of children(row, 'tc')) {
			if (!only(cell, ['tcPr', 'p']) || !children(cell, 'p').length) return false;
			const props = first(cell, 'tcPr');
			if (props && ['gridSpan', 'vMerge', 'hMerge', 'tcPrChange'].some((key) => first(props, key)))
				return false;
			for (const paragraph of children(cell, 'p')) {
				if (!only(paragraph, ['pPr', 'r'])) return false;
				for (const run of children(paragraph, 'r'))
					if (!only(run, ['rPr', 't', 'tab', 'br', 'cr', 'noBreakHyphen'])) return false;
			}
		}
	}
	return true;
}

const layout = (table: Table) =>
	table.rows.map((row) => row.map((cell) => cell.paragraphs.map((p) => p.id)));

/** Table/cell descriptive fields this writer does not yet serialize back to XML. */
const TABLE_DESCRIPTOR_KEYS = [
	'grid',
	'widthTwips',
	'alignment',
	'indentTwips',
	'borders',
	'style',
	'look',
] as const;
const CELL_DESCRIPTOR_KEYS = [
	'gridSpan',
	'verticalMerge',
	'widthTwips',
	'verticalAlign',
	'shadingFill',
	'shadingThemeFill',
	'borders',
	'margins',
] as const;
function pick(source: object, keys: readonly string[]): Record<string, unknown> {
	const record = source as Record<string, unknown>;
	const result: Record<string, unknown> = {};
	for (const key of keys) if (record[key] !== undefined) result[key] = record[key];
	return result;
}
/**
 * Table grid/width/border/style and per-cell width/merge/shading/border/margin values render
 * (see resolve-table.ts) but are not yet serialized on save; reject rather than silently drop edits.
 */
function assertNoDescriptorEdits(table: Table, base: Table | undefined): void {
	if (!base) return;
	if (
		JSON.stringify(pick(table, TABLE_DESCRIPTOR_KEYS)) !==
		JSON.stringify(pick(base, TABLE_DESCRIPTOR_KEYS))
	)
		throw new Error(
			'Cannot edit table grid widths, width, alignment, indent, borders, or style; only cell text and row/column structure changes are supported. The original DOCX package remains unchanged.',
		);
	const baseCellById = new Map<string, TableCell>();
	for (const row of base.rows)
		for (const cell of row) {
			const id = cell.paragraphs[0]?.id;
			if (id) baseCellById.set(id, cell);
		}
	for (const row of table.rows)
		for (const cell of row) {
			const id = cell.paragraphs[0]?.id;
			const source = id ? baseCellById.get(id) : undefined;
			if (!source) continue;
			if (
				JSON.stringify(pick(cell, CELL_DESCRIPTOR_KEYS)) !==
				JSON.stringify(pick(source, CELL_DESCRIPTOR_KEYS))
			)
				throw new Error(
					'Cannot edit table cell width, merge, vertical alignment, shading, borders, or margins on an existing cell; only cell text is supported. The original DOCX package remains unchanged.',
				);
		}
}

export function writeTable(
	doc: XmlDocument,
	table: Table,
	node: XmlElement,
	base: Table | undefined,
	writeParagraph: ParagraphWriter,
	replaceSlots: SlotsWriter,
): XmlElement {
	if (base && JSON.stringify(table) === JSON.stringify(base)) return node;
	assertNoDescriptorEdits(table, base);
	const structural = base && JSON.stringify(layout(table)) !== JSON.stringify(layout(base));
	if (structural && !canEditTableStructure(node))
		throw new Error(
			'Cannot change the structure of a merged, nested, or complex table. The original DOCX package remains unchanged.',
		);
	const width = table.rows[0]?.length ?? 0;
	if (
		!width ||
		table.rows.some((row) => row.length !== width || row.some((cell) => !cell.paragraphs.length))
	)
		throw new Error('Tables must contain rectangular rows and at least one paragraph per cell.');
	const oldRows = children(node, 'tr');
	type Source = { row: number; col: number; cell: XmlElement; paragraphs: Paragraph[] };
	const sourceById = new Map<string, Source>();
	base?.rows.forEach((row, ri) =>
		row.forEach((cell, ci) => {
			const oldCell = children(expectDefined(oldRows[ri], `table row ${ri}`), 'tc')[ci];
			if (oldCell)
				for (const paragraph of cell.paragraphs)
					sourceById.set(paragraph.id, {
						row: ri,
						col: ci,
						cell: oldCell,
						paragraphs: cell.paragraphs,
					});
		}),
	);
	const usedCells = new Set<XmlElement>();
	const usedRows = new Set<number>();
	const columns = new Map<number, number>();
	const nextRows = table.rows.map((row) => {
		const sources = row.map((cell) =>
			cell.paragraphs.map((p) => sourceById.get(p.id)).find(Boolean),
		);
		const sourceRow = sources.find(Boolean)?.row;
		if (sources.some((source) => source && source.row !== sourceRow))
			throw new Error('Moving paragraphs between table rows is not supported.');
		if (sourceRow !== undefined && usedRows.has(sourceRow))
			throw new Error('Splitting an existing table row is not supported.');
		if (sourceRow !== undefined) usedRows.add(sourceRow);
		const tr =
			sourceRow === undefined
				? makeW(doc, 'tr')
				: expectDefined(oldRows[sourceRow], `table row ${sourceRow}`);
		const oldCells = children(tr, 'tc');
		const nextCells = row.map((cell, ci) => {
			const source = sources[ci];
			if (source && usedCells.has(source.cell))
				throw new Error('Splitting an existing table cell is not supported.');
			if (source) {
				usedCells.add(source.cell);
				if (columns.has(ci) && columns.get(ci) !== source.col)
					throw new Error('Inconsistent table column edits are not supported.');
				columns.set(ci, source.col);
			}
			if (
				cell.paragraphs.some(
					(p) => sourceById.has(p.id) && sourceById.get(p.id)?.cell !== source?.cell,
				)
			)
				throw new Error('Moving paragraphs between table cells is not supported.');
			const tc = source?.cell ?? makeW(doc, 'tc');
			const oldParagraphs = children(tc, 'p');
			const nextParagraphs = cell.paragraphs.map((paragraph) => {
				const index = source?.paragraphs.findIndex((p) => p.id === paragraph.id) ?? -1;
				return writeParagraph(
					doc,
					paragraph,
					oldParagraphs[index] ?? makeW(doc, 'p'),
					source?.paragraphs[index],
				);
			});
			replaceSlots(doc, tc, oldParagraphs, nextParagraphs);
			return tc;
		});
		replaceSlots(doc, tr, oldCells, nextCells);
		return tr;
	});
	replaceSlots(doc, node, oldRows, nextRows);
	const grid = first(node, 'tblGrid');
	if (
		structural &&
		grid &&
		(width !== base.rows[0]?.length || [...columns].some(([next, old]) => next !== old))
	) {
		const oldColumns = children(grid, 'gridCol');
		const nextColumns = Array.from({ length: width }, (_, index) => {
			const old = oldColumns[columns.get(index) ?? -1];
			return (
				old ??
				(oldColumns[Math.min(index, oldColumns.length - 1)]?.cloneNode(true) as
					| XmlElement
					| undefined) ??
				makeW(doc, 'gridCol')
			);
		});
		replaceSlots(doc, grid, oldColumns, nextColumns);
	}
	return node;
}
