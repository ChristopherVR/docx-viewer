import { applyCellFill } from './table-shading';
import { closeHistory } from 'prosemirror-history';
import type { EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';

/** Borders for the cells of a simple (structure-editable) table, edited per cell like Word's Borders menu. */
export type CellSide = 'top' | 'bottom' | 'left' | 'right';
export type TableBorderPreset =
	| 'none'
	| 'bottom'
	| 'top'
	| 'left'
	| 'right'
	| 'all'
	| 'outside'
	| 'insideH'
	| 'insideV';
export interface BorderPen {
	style: string;
	sizeEighthPoints: number;
	color?: string;
}
type StoredSide = { style?: string; sizeEighthPoints?: number; color?: string };
type StoredBorders = Partial<Record<CellSide, StoredSide>>;
export interface TableBorderContext {
	tablePos: number;
	rows: number;
	columns: number;
	/** Inclusive cell rectangle covered by the selection. */
	rect: { top: number; bottom: number; left: number; right: number };
	cells: Array<{ pos: number; row: number; column: number; node: ProseMirrorNode }>;
}

/** Word's default border pen: a single 0.5 pt automatic-colour line. */
export const DEFAULT_PEN: BorderPen = { style: 'single', sizeEighthPoints: 4 };
const OPPOSITE: Record<CellSide, CellSide> = {
	top: 'bottom',
	bottom: 'top',
	left: 'right',
	right: 'left',
};

const parse = (value: unknown): StoredBorders => {
	if (typeof value !== 'string') return {};
	try {
		const parsed = JSON.parse(value) as unknown;
		return parsed && typeof parsed === 'object' ? (parsed as StoredBorders) : {};
	} catch {
		return {};
	}
};
const visible = (side: StoredSide | undefined) =>
	Boolean(side && side.style && side.style !== 'none' && side.style !== 'nil');

/** The table and the cells the selection touches, or null outside a simple table. */
export function tableBorderContext(state: EditorState): TableBorderContext | null {
	const { $from, from, to } = state.selection;
	for (let depth = $from.depth; depth > 0; depth--) {
		const table = $from.node(depth);
		if (table.type.name !== 'table') continue;
		if (table.attrs.structureEditable === false) return null;
		const tablePos = $from.before(depth);
		const cells: TableBorderContext['cells'] = [];
		table.forEach((row, rowOffset, rowIndex) => {
			row.forEach((cell, cellOffset, column) => {
				const pos = tablePos + 1 + rowOffset + 1 + cellOffset;
				if (Math.max(from, pos + 1) <= Math.min(to, pos + cell.nodeSize - 1))
					cells.push({ pos, row: rowIndex, column, node: cell });
			});
		});
		if (!cells.length) return null;
		return {
			tablePos,
			rows: table.childCount,
			columns: table.firstChild?.childCount ?? 0,
			rect: {
				top: Math.min(...cells.map((cell) => cell.row)),
				bottom: Math.max(...cells.map((cell) => cell.row)),
				left: Math.min(...cells.map((cell) => cell.column)),
				right: Math.max(...cells.map((cell) => cell.column)),
			},
			cells,
		};
	}
	return null;
}

type Edit = { row: number; column: number; side: CellSide; pen: BorderPen | null };
type Rect = TableBorderContext['rect'];

/** The cell edges a preset covers inside `rect`. */
function edgesOf(preset: Exclude<TableBorderPreset, 'none'>, rect: Rect) {
	const edges: Array<{ row: number; column: number; side: CellSide }> = [];
	for (let row = rect.top; row <= rect.bottom; row++)
		for (let column = rect.left; column <= rect.right; column++) {
			const outside = {
				top: row === rect.top,
				bottom: row === rect.bottom,
				left: column === rect.left,
				right: column === rect.right,
			};
			for (const side of ['top', 'bottom', 'left', 'right'] as const) {
				const isOutside = outside[side];
				const horizontal = side === 'top' || side === 'bottom';
				const wanted =
					preset === 'all' ||
					(preset === 'outside' && isOutside) ||
					(preset === 'insideH' && horizontal && !isOutside) ||
					(preset === 'insideV' && !horizontal && !isOutside) ||
					(preset === side && isOutside);
				if (wanted) edges.push({ row, column, side });
			}
		}
	return edges;
}

const current = (ctx: TableBorderContext, state: EditorState, row: number, column: number) => {
	const cell = state.doc.nodeAt(ctx.tablePos)?.child(row).child(column);
	return parse(cell?.attrs.borders);
};

/** Writes each edit to its cell and to the neighbour that shares the edge. */
function applyEdits(
	view: EditorView,
	ctx: TableBorderContext,
	edits: Edit[],
	shading?: { scope: 'cell' | 'table'; fill: string | null },
): boolean {
	const all = new Map<string, Edit>();
	const add = (edit: Edit) => all.set(`${edit.row}:${edit.column}:${edit.side}`, edit);
	for (const edit of edits) {
		add(edit);
		const shared = {
			row: edit.row + (edit.side === 'bottom' ? 1 : edit.side === 'top' ? -1 : 0),
			column: edit.column + (edit.side === 'right' ? 1 : edit.side === 'left' ? -1 : 0),
		};
		if (
			shared.row >= 0 &&
			shared.row < ctx.rows &&
			shared.column >= 0 &&
			shared.column < ctx.columns
		)
			add({ ...shared, side: OPPOSITE[edit.side], pen: edit.pen });
	}
	const tr = view.state.tr;
	const table = view.state.doc.nodeAt(ctx.tablePos);
	if (!table) return false;
	const grouped = new Map<string, Edit[]>();
	for (const edit of all.values())
		grouped.set(`${edit.row}:${edit.column}`, [
			...(grouped.get(`${edit.row}:${edit.column}`) ?? []),
			edit,
		]);
	let rowOffset = ctx.tablePos + 1;
	table.forEach((row, _offset, rowIndex) => {
		let cellPos = rowOffset + 1;
		row.forEach((cell, _cellOffset, column) => {
			const group = grouped.get(`${rowIndex}:${column}`);
			if (group) {
				const visual = parse(cell.attrs.borders);
				const direct = parse(cell.attrs.directBorders);
				for (const edit of group) {
					const stored: StoredSide = edit.pen
						? {
								style: edit.pen.style,
								sizeEighthPoints: edit.pen.sizeEighthPoints,
								...(edit.pen.color ? { color: edit.pen.color } : {}),
							}
						: { style: 'none', sizeEighthPoints: 4 };
					visual[edit.side] = stored;
					direct[edit.side] = stored;
				}
				tr.setNodeMarkup(cellPos, undefined, {
					...cell.attrs,
					borders: JSON.stringify(visual),
					directBorders: JSON.stringify(direct),
				});
			}
			cellPos += cell.nodeSize;
		});
		rowOffset += row.nodeSize;
	});
	if (shading) applyCellFill(tr, ctx, shading.scope, shading.fill);
	if (!tr.docChanged) return false;
	view.dispatch(closeHistory(tr));
	return true;
}

/**
 * Home > Borders for table cells. A side preset toggles: when every targeted edge already shows a
 * line they are removed, otherwise they are drawn with `pen`. `none` clears every edge of the
 * selection, inside and out.
 */
export function setCellBorders(
	view: EditorView,
	preset: TableBorderPreset,
	pen: BorderPen = DEFAULT_PEN,
): boolean {
	const ctx = tableBorderContext(view.state);
	if (!view.editable || !ctx) return false;
	const edges = edgesOf(preset === 'none' ? 'all' : preset, ctx.rect);
	if (!edges.length) return false;
	const allHave =
		preset !== 'none' &&
		edges.every((edge) => visible(current(ctx, view.state, edge.row, edge.column)[edge.side]));
	return applyEdits(
		view,
		ctx,
		edges.map((edge) => ({ ...edge, pen: preset === 'none' || allHave ? null : pen })),
	);
}

export interface CellBorderSettings {
	scope: 'cell' | 'table';
	/** Edges to draw with `pen`; the others of the scope are cleared. */
	sides: Partial<Record<CellSide | 'insideH' | 'insideV', boolean>>;
	pen: BorderPen;
	/** Omitted preserves existing fill, null explicitly clears it. */
	fill?: string | null;
}

const wholeTable = (ctx: TableBorderContext): Rect => ({
	top: 0,
	bottom: ctx.rows - 1,
	left: 0,
	right: ctx.columns - 1,
});

/** Borders and Shading dialog, Cell or Table scope: draws the checked edges and clears the others. */
export function applyCellBorderSettings(view: EditorView, settings: CellBorderSettings): boolean {
	const ctx = tableBorderContext(view.state);
	if (!view.editable || !ctx) return false;
	const rect = settings.scope === 'table' ? wholeTable(ctx) : ctx.rect;
	const edits: Edit[] = [];
	for (const [key, preset] of [
		['top', 'top'],
		['bottom', 'bottom'],
		['left', 'left'],
		['right', 'right'],
		['insideH', 'insideH'],
		['insideV', 'insideV'],
	] as const)
		for (const edge of edgesOf(preset, rect))
			edits.push({ ...edge, pen: settings.sides[key] ? settings.pen : null });
	return applyEdits(
		view,
		ctx,
		edits,
		settings.fill === undefined ? undefined : { scope: settings.scope, fill: settings.fill },
	);
}

/** What the dialog shows for `scope`: which edges have a line, and the first line's pen. */
export function readCellBorderSettings(
	state: EditorState,
	scope: 'cell' | 'table',
): CellBorderSettings | null {
	const ctx = tableBorderContext(state);
	if (!ctx) return null;
	const rect = scope === 'table' ? wholeTable(ctx) : ctx.rect;
	const sides: CellBorderSettings['sides'] = {};
	let pen: BorderPen | undefined;
	for (const key of ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'] as const) {
		const edges = edgesOf(key, rect);
		sides[key] =
			edges.length > 0 &&
			edges.every((edge) => {
				const side = current(ctx, state, edge.row, edge.column)[edge.side];
				if (visible(side) && !pen)
					pen = {
						style: side?.style ?? 'single',
						sizeEighthPoints: side?.sizeEighthPoints ?? 4,
						...(side?.color ? { color: side.color } : {}),
					};
				return visible(side);
			});
	}
	return { scope, sides, pen: pen ?? DEFAULT_PEN };
}
