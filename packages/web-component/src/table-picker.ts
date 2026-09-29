import { localeOf, translate, translateTemplate } from './localization';
import { MAX_TABLE_COLUMNS as MAX_COLUMNS, MAX_TABLE_ROWS as MAX_ROWS } from './ribbon-commands';
import { mountPopover } from './ribbon-popover';

/** Word's table size picker: how many columns and rows fit the grid, and the hard limits. */
export const PICKER_COLUMNS = 10;
export const PICKER_ROWS = 8;

/** Clamps a typed size to what the editor inserts. */
export function clampTableSize(value: number, max: number): number {
	return Number.isFinite(value) ? Math.min(max, Math.max(1, Math.floor(value))) : 1;
}

/**
 * The Insert > Table drop-down: hover a grid to preview "columns x rows" and click to insert, or
 * type a size. Arrow keys move over the grid and Enter inserts, so it works without a mouse.
 */
export function openTablePicker(
	anchor: HTMLElement,
	onPick: (rows: number, columns: number) => void,
	onClose: () => void,
): void {
	const locale = localeOf(anchor);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover table-picker';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', translate(locale, 'Insert table'));
	const caption = document.createElement('div');
	caption.className = 'table-picker-caption';
	const grid = document.createElement('div');
	grid.className = 'table-picker-grid';
	grid.setAttribute('role', 'grid');
	grid.tabIndex = 0;
	grid.setAttribute('aria-label', translate(locale, 'Table size'));
	const cells: HTMLElement[] = [];
	let rows = 2;
	let columns = 2;
	const paint = () => {
		caption.textContent = translateTemplate(locale, 'table.size', { columns, rows });
		cells.forEach((cell, index) => {
			const row = Math.floor(index / PICKER_COLUMNS);
			const column = index % PICKER_COLUMNS;
			cell.classList.toggle('active', row < rows && column < columns);
		});
	};
	for (let row = 0; row < PICKER_ROWS; row++)
		for (let column = 0; column < PICKER_COLUMNS; column++) {
			const cell = document.createElement('div');
			cell.setAttribute('role', 'gridcell');
			cell.dataset.rows = String(row + 1);
			cell.dataset.columns = String(column + 1);
			cell.addEventListener('mouseenter', () => {
				rows = row + 1;
				columns = column + 1;
				paint();
			});
			cell.addEventListener('mousedown', (event) => event.preventDefault());
			cell.addEventListener('click', () => {
				onPick(row + 1, column + 1);
				onClose();
			});
			cells.push(cell);
			grid.append(cell);
		}
	grid.addEventListener('keydown', (event) => {
		const step: Record<string, [number, number]> = {
			ArrowRight: [0, 1],
			ArrowLeft: [0, -1],
			ArrowDown: [1, 0],
			ArrowUp: [-1, 0],
		};
		const move = step[event.key];
		if (move) {
			event.preventDefault();
			rows = Math.min(PICKER_ROWS, Math.max(1, rows + move[0]));
			columns = Math.min(PICKER_COLUMNS, Math.max(1, columns + move[1]));
			paint();
		} else if (event.key === 'Enter') {
			event.preventDefault();
			onPick(rows, columns);
			onClose();
		}
	});
	const field = (label: string, value: number, max: number) => {
		const wrap = document.createElement('label');
		const text = document.createElement('span');
		text.textContent = translate(locale, label as never);
		const input = document.createElement('input');
		input.type = 'number';
		input.min = '1';
		input.max = String(max);
		input.value = String(value);
		input.setAttribute('aria-label', translate(locale, label as never));
		wrap.append(text, input);
		return { wrap, input };
	};
	const columnField = field('Columns', 5, MAX_COLUMNS);
	const rowField = field('Rows', 2, MAX_ROWS);
	const insert = document.createElement('button');
	insert.type = 'button';
	insert.className = 'table-picker-insert';
	insert.textContent = translate(locale, 'Insert table');
	insert.addEventListener('click', () => {
		onPick(
			clampTableSize(Number(rowField.input.value), MAX_ROWS),
			clampTableSize(Number(columnField.input.value), MAX_COLUMNS),
		);
		onClose();
	});
	const custom = document.createElement('div');
	custom.className = 'table-picker-custom';
	custom.append(columnField.wrap, rowField.wrap, insert);
	pop.append(caption, grid, custom);
	paint();
	if (mountPopover(anchor, pop, anchor)) grid.focus();
}
