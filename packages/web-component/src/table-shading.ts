import type { EditorState, Transaction } from 'prosemirror-state';
import type { TableBorderContext } from './table-border-commands';

/** Undefined means mixed fills; null means no visible fill. */
export function readCellFill(
	ctx: TableBorderContext,
	state: EditorState,
	scope: 'cell' | 'table',
): string | null | undefined {
	const fills = new Set<string | null>();
	const table = state.doc.nodeAt(ctx.tablePos);
	table?.forEach((row, _offset, rowIndex) =>
		row.forEach((cell, _cellOffset, column) => {
			if (
				scope === 'cell' &&
				(rowIndex < ctx.rect.top ||
					rowIndex > ctx.rect.bottom ||
					column < ctx.rect.left ||
					column > ctx.rect.right)
			)
				return;
			const fill = cell.attrs.shadingFill;
			fills.add(
				typeof fill === 'string' && /^#[0-9a-f]{6}$/i.test(fill) ? fill.toLowerCase() : null,
			);
		}),
	);
	return fills.size === 1 ? [...fills][0] : undefined;
}

/** Adds shading to the border transaction, so applying the dialog is one undo step. */
export function applyCellFill(
	tr: Transaction,
	ctx: TableBorderContext,
	scope: 'cell' | 'table',
	fill: string | null,
): void {
	const table = tr.doc.nodeAt(ctx.tablePos);
	if (!table) return;
	table.forEach((row, rowOffset, rowIndex) =>
		row.forEach((cell, cellOffset, column) => {
			if (
				scope === 'cell' &&
				(rowIndex < ctx.rect.top ||
					rowIndex > ctx.rect.bottom ||
					column < ctx.rect.left ||
					column > ctx.rect.right)
			)
				return;
			tr.setNodeMarkup(ctx.tablePos + 2 + rowOffset + cellOffset, undefined, {
				...cell.attrs,
				shadingFill: fill,
				directShadingFill: fill ?? 'auto',
				shadingEdited: true,
			});
		}),
	);
}
