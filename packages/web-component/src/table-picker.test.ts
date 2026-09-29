// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { insertTable } from './ribbon-commands';
import { createRibbon } from './ribbon';
import { schema } from './schema';
import { clampTableSize, openTablePicker } from './table-picker';

afterEach(() => (document.body.innerHTML = ''));

function open() {
	const anchor = document.createElement('button');
	document.body.append(anchor);
	const picks: Array<[number, number]> = [];
	let closed = 0;
	openTablePicker(
		anchor,
		(rows, columns) => picks.push([rows, columns]),
		() => closed++,
	);
	return { picks, closed: () => closed, pop: document.querySelector('.table-picker')! };
}
const cell = (root: ParentNode, rows: number, columns: number) =>
	root.querySelector<HTMLElement>(
		`[role="gridcell"][data-rows="${rows}"][data-columns="${columns}"]`,
	)!;

describe('table picker', () => {
	it('previews the size on hover and inserts it on click', () => {
		const { pop, picks, closed } = open();
		expect(pop.querySelector('.table-picker-caption')!.textContent).toBe('2 × 2 Table');
		cell(pop, 3, 5).dispatchEvent(new MouseEvent('mouseenter'));
		expect(pop.querySelector('.table-picker-caption')!.textContent).toBe('5 × 3 Table');
		expect(pop.querySelectorAll('.active')).toHaveLength(15);
		cell(pop, 3, 5).click();
		expect(picks).toEqual([[3, 5]]);
		expect(closed()).toBe(1);
	});

	it('moves with the arrow keys and inserts with Enter', () => {
		const { pop, picks } = open();
		const grid = pop.querySelector<HTMLElement>('[role="grid"]')!;
		for (const key of ['ArrowRight', 'ArrowRight', 'ArrowDown'])
			grid.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
		grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
		expect(picks).toEqual([[3, 4]]);
	});

	it('does not move past the grid edges', () => {
		const { pop } = open();
		const grid = pop.querySelector<HTMLElement>('[role="grid"]')!;
		for (let i = 0; i < 20; i++)
			grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowLeft', bubbles: true }));
		expect(pop.querySelector('.table-picker-caption')!.textContent).toBe('1 × 2 Table');
		for (let i = 0; i < 20; i++)
			grid.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(pop.querySelector('.table-picker-caption')!.textContent).toBe('1 × 8 Table');
	});

	it('inserts a typed size, clamped to the limits', () => {
		const { pop, picks } = open();
		const columns = pop.querySelector<HTMLInputElement>('input[aria-label="Columns"]')!;
		const rows = pop.querySelector<HTMLInputElement>('input[aria-label="Rows"]')!;
		columns.value = '500';
		rows.value = '7';
		pop.querySelector<HTMLButtonElement>('.table-picker-insert')!.click();
		expect(picks).toEqual([[7, 20]]);
		expect(clampTableSize(Number.NaN, 20)).toBe(1);
		expect(clampTableSize(0, 20)).toBe(1);
		expect(clampTableSize(3.9, 20)).toBe(3);
	});

	it('opens from the Insert table button', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const ribbon = createRibbon();
		host.append(ribbon);
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('[aria-label="Insert table"]')!.click();
		expect(seen).toEqual([{ type: 'tablePicker' }]);
	});
});

describe('insertTable size', () => {
	const editor = () =>
		new EditorView(document.createElement('div'), {
			state: EditorState.create({
				doc: schema.node('doc', null, [schema.nodes.paragraph!.create({ id: 'p' })]),
				schema,
			}),
		});
	const dims = (view: EditorView) => {
		let found = [0, 0];
		view.state.doc.descendants((node) => {
			if (node.type.name === 'table') found = [node.childCount, node.firstChild!.childCount];
		});
		return found;
	};

	it('builds the requested rows and columns, 2 x 2 by default', () => {
		const sized = editor();
		insertTable(sized, undefined, { rows: 4, columns: 3 });
		expect(dims(sized)).toEqual([4, 3]);
		const plain = editor();
		insertTable(plain);
		expect(dims(plain)).toEqual([2, 2]);
	});

	it('caps absurd sizes', () => {
		const view = editor();
		insertTable(view, undefined, { rows: 100000, columns: 100000 });
		expect(dims(view)).toEqual([100, 20]);
	});
});
