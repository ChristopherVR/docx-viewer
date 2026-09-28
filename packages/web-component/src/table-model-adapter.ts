// Table <-> ProseMirror conversion. Merged/nested/complex tables (structureEditable === false)
// preserve their full source row/cell/descriptor structure and only patch paragraph content by
// id, because vertical-merge continuation cells and nested-table previews have no 1:1 visible
// ProseMirror node to rebuild from; simple tables keep the existing rebuild-from-doc path.
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Paragraph, Table, TableCell, TableStyleCatalog } from '@christophervr/docx-core';
import { schema } from './schema';
import { resolveCellVisuals } from './table-visuals';

function cellKey(cell: TableCell): string {
	return cell.paragraphs[0]?.id ?? '';
}

export function tableNode(
	block: Table,
	paragraphNode: (paragraph: Paragraph) => ProseMirrorNode,
	tableStyles?: TableStyleCatalog,
): ProseMirrorNode {
	const columnCount = Math.max(
		1,
		...block.rows.map((row) => row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0)),
	);
	const rows = block.rows.map((row, rowIndex) => {
		const cells: ProseMirrorNode[] = [];
		let gridColumn = 0;
		row.forEach((cell, columnIndex) => {
			const column = gridColumn;
			gridColumn += cell.gridSpan ?? 1;
			if (cell.verticalMerge === 'continue') return;
			let rowspan = 1;
			if (cell.verticalMerge === 'restart') {
				let next = rowIndex + 1;
				while (block.rows[next]?.[columnIndex]?.verticalMerge === 'continue') {
					rowspan++;
					next++;
				}
			}
			const visuals = resolveCellVisuals(
				block,
				cell,
				{
					row: rowIndex,
					lastRow: rowIndex + rowspan - 1,
					column,
					lastColumn: column + (cell.gridSpan ?? 1) - 1,
					rowCount: block.rows.length,
					columnCount,
				},
				tableStyles,
			);
			const children = cell.paragraphs.map(paragraphNode);
			for (const preview of cell.nestedTables ?? [])
				children.push(
					schema.nodes.nestedTablePreview.create({ rowsJson: JSON.stringify(preview.rows) }),
				);
			cells.push(
				schema.node(
					'tableCell',
					{
						sourceCellKey: cellKey(cell),
						colspan: cell.gridSpan ?? 1,
						rowspan,
						widthTwips: cell.widthTwips ?? null,
						verticalAlign: cell.verticalAlign ?? null,
						shadingFill: visuals.shadingFill ?? null,
						borders: visuals.borders ? JSON.stringify(visuals.borders) : null,
						margins: cell.margins ? JSON.stringify(cell.margins) : null,
					},
					children,
				),
			);
		});
		const heightTwips = block.rowProperties?.[rowIndex]?.heightTwips;
		return schema.node('tableRow', heightTwips ? { heightPx: heightTwips / 15 } : null, cells);
	});
	return schema.node(
		'table',
		{
			id: block.id,
			structureEditable: block.structureEditable !== false,
			widthTwips: block.widthTwips ?? null,
			alignment: block.alignment ?? null,
			indentTwips: block.indentTwips ?? null,
			borders: block.borders ? JSON.stringify(block.borders) : null,
			cellMargins: block.cellMargins ? JSON.stringify(block.cellMargins) : null,
		},
		rows,
	);
}

/** Rebuilds a simple (structurally editable) table directly from its visible ProseMirror rows. */
export function convertSimpleTable(
	node: ProseMirrorNode,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
	prior?: Table,
): Table['rows'] {
	// Cell width, shading, borders and margins have no editor controls; keep each existing cell's
	// values (matched by its source cell key) so editing text never reads as a formatting change.
	const priorCells = new Map<string, TableCell>();
	for (const row of prior?.rows ?? []) for (const cell of row) priorCells.set(cellKey(cell), cell);
	const rows: Table['rows'] = [];
	node.forEach((row) => {
		const cells: Table['rows'][number] = [];
		row.forEach((cell) => {
			const paragraphs: Paragraph[] = [];
			cell.forEach((paragraph) => {
				if (paragraph.type.name === 'paragraph') paragraphs.push(convertParagraph(paragraph));
			});
			const source = priorCells.get(String(cell.attrs.sourceCellKey || ''));
			const { paragraphs: _previous, ...properties } = source ?? { paragraphs: [] };
			cells.push({ ...structuredClone(properties), paragraphs });
		});
		rows.push(cells);
	});
	return rows;
}

/**
 * Patches paragraph content back into a merged/nested/complex table's prior structure by cell
 * identity, leaving row/cell layout, merges and descriptive properties exactly as they were.
 */
export function convertMergedTable(
	node: ProseMirrorNode,
	prior: Table,
	convertParagraph: (paragraph: ProseMirrorNode) => Paragraph,
): Table {
	const updated = structuredClone(prior);
	const cellByKey = new Map<string, TableCell>();
	for (const row of updated.rows) for (const cell of row) cellByKey.set(cellKey(cell), cell);
	node.forEach((row) =>
		row.forEach((cellNode) => {
			const target = cellByKey.get(String(cellNode.attrs.sourceCellKey || ''));
			if (!target) return;
			const paragraphs: Paragraph[] = [];
			cellNode.forEach((child) => {
				if (child.type.name === 'paragraph') paragraphs.push(convertParagraph(child));
			});
			if (paragraphs.length) target.paragraphs = paragraphs;
		}),
	);
	return updated;
}
