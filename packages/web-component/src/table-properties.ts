import type { Table, TableCell, TableCellMargins } from 'docx-core';
import { closeHistory } from 'prosemirror-history';
import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type TableRowProperties = NonNullable<Table['rowProperties']>[number];
export type RowPatch = { [K in keyof TableRowProperties]?: TableRowProperties[K] | undefined };

/** Only rectangular tables currently expose formatting controls. */
export function tablePropertiesContext(state: EditorState) {
	const { $from, $to } = state.selection;
	for (let depth = $from.depth; depth > 0; depth--) {
		const table = $from.node(depth);
		if (table.type.name !== 'table') continue;
		if (
			table.attrs.structureEditable === false ||
			$to.depth <= depth ||
			$to.node(depth) !== table ||
			$from.depth <= depth
		)
			return null;
		const rows: Array<{ pos: number; properties: TableRowProperties }> = [];
		const cells: Array<{
			pos: number;
			margins: TableCellMargins;
			verticalAlign: TableCell['verticalAlign'];
		}> = [];
		let pos = $from.start(depth);
		table.forEach((row, offset, index) => {
			if (index >= $from.index(depth) && index <= $to.index(depth))
				rows.push({ pos: pos + offset, properties: row.attrs.properties ?? {} });
			row.forEach((cell, cellOffset) => {
				const cellPos = pos + offset + 1 + cellOffset;
				if (
					Math.max(state.selection.from, cellPos + 1) <=
					Math.min(state.selection.to, cellPos + cell.nodeSize - 1)
				)
					cells.push({
						pos: cellPos,
						margins: cell.attrs.margins ? (JSON.parse(cell.attrs.margins) as TableCellMargins) : {},
						verticalAlign: (cell.attrs.verticalAlign ?? undefined) as TableCell['verticalAlign'],
					});
			});
		});
		return {
			pos: $from.before(depth),
			table,
			rows,
			cells,
			margins: (table.attrs.cellMargins
				? JSON.parse(table.attrs.cellMargins)
				: {}) as TableCellMargins,
		};
	}
	return null;
}

/** Patches only fields changed in the dialog; all edits form one history step. */
export function applyTableProperties(
	view: EditorView,
	rowPatch: RowPatch,
	marginPatch: TableCellMargins,
	/** Null removes all direct cell overrides; omitted leaves them untouched. */
	cellMarginPatch?: TableCellMargins | null,
	/** Null removes the direct alignment override; omitted leaves it untouched. */
	cellAlignment?: TableCell['verticalAlign'] | null,
): boolean {
	const context = tablePropertiesContext(view.state);
	if (!view.editable || !context) return false;
	if (
		!Object.keys(rowPatch).length &&
		!Object.keys(marginPatch).length &&
		cellMarginPatch === undefined &&
		cellAlignment === undefined
	)
		return true;
	const tr = closeHistory(view.state.tr);
	for (const row of context.rows) {
		const node = tr.doc.nodeAt(row.pos)!;
		const properties = { ...row.properties, ...rowPatch };
		for (const key of Object.keys(properties) as Array<keyof TableRowProperties>)
			if (properties[key] === undefined) delete properties[key];
		tr.setNodeMarkup(row.pos, undefined, {
			...node.attrs,
			properties,
			heightPx: properties.heightTwips ? properties.heightTwips / 15 : null,
		});
	}
	if (Object.keys(marginPatch).length)
		tr.setNodeMarkup(context.pos, undefined, {
			...context.table.attrs,
			cellMargins: JSON.stringify({ ...context.margins, ...marginPatch }),
		});
	if (cellMarginPatch !== undefined || cellAlignment !== undefined)
		for (const cell of context.cells) {
			const node = tr.doc.nodeAt(cell.pos)!;
			tr.setNodeMarkup(cell.pos, undefined, {
				...node.attrs,
				...(cellMarginPatch === undefined
					? {}
					: {
							margins:
								cellMarginPatch === null
									? null
									: JSON.stringify({ ...cell.margins, ...cellMarginPatch }),
						}),
				...(cellAlignment === undefined ? {} : { verticalAlign: cellAlignment }),
			});
		}
	view.dispatch(tr);
	return true;
}
