// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { EditorView } from 'prosemirror-view';
import { executeTableCommand, type TableCommand } from './table-commands';
import { schema } from './schema';

function makeTable(rows = 2, columns = 2) {
	return schema.nodes.table.create(
		{ id: 'grid' },
		Array.from({ length: rows }, (_, row) =>
			schema.nodes.tableRow.create(
				null,
				Array.from({ length: columns }, (_, column) =>
					schema.nodes.tableCell.create(
						null,
						schema.nodes.paragraph.create({ id: `cell-${row}-${column}` }, [
							schema.text(`r${row}c${column}`),
						]),
					),
				),
			),
		),
	);
}

function setup(table = makeTable(), selectedCell = 'cell-0-0') {
	const doc = schema.nodes.doc.create(null, [
		schema.nodes.paragraph.create({ id: 'before' }, [schema.text('before')]),
		table,
		schema.nodes.paragraph.create({ id: 'after' }, [schema.text('after')]),
	]);
	let cellPosition = -1;
	doc.descendants((node, pos) => {
		if (node.type === schema.nodes.paragraph && node.attrs.id === selectedCell) {
			cellPosition = pos + 1;
			return false;
		}
		return true;
	});
	const state = EditorState.create({
		doc,
		selection: TextSelection.create(doc, cellPosition),
		plugins: [history()],
	});
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, { state });
	return { host, view };
}

function grid(view: EditorView): string[][] {
	const table: string[][] = [];
	view.state.doc.descendants((node) => {
		if (node.type !== schema.nodes.table) return;
		for (let row = 0; row < node.childCount; row++) {
			const values: string[] = [];
			for (let column = 0; column < node.child(row).childCount; column++)
				values.push(node.child(row).child(column).textContent);
			table.push(values);
		}
	});
	return table;
}

function run(view: EditorView, action: TableCommand) {
	return executeTableCommand(view, action);
}

describe('table structural commands', () => {
	afterEach(() => document.body.replaceChildren());
	it('uses the row index when the selected column differs from the selected row', () => {
		const { view } = setup(makeTable(3, 2), 'cell-2-0');
		expect(run(view, 'deleteRow')).toBe(true);
		expect(grid(view)).toEqual([
			['r0c0', 'r0c1'],
			['r1c0', 'r1c1'],
		]);
		view.destroy();
	});

	it.each([
		[
			'rowBefore',
			[
				['', ''],
				['r0c0', 'r0c1'],
				['r1c0', 'r1c1'],
			],
		],
		[
			'rowAfter',
			[
				['r0c0', 'r0c1'],
				['', ''],
				['r1c0', 'r1c1'],
			],
		],
		[
			'columnBefore',
			[
				['', 'r0c0', 'r0c1'],
				['', 'r1c0', 'r1c1'],
			],
		],
		[
			'columnAfter',
			[
				['r0c0', '', 'r0c1'],
				['r1c0', '', 'r1c1'],
			],
		],
	] as const)('adds cells for %s and keeps old cell contents', (action, expected) => {
		const { view } = setup();
		expect(run(view, action)).toBe(true);
		expect(grid(view)).toEqual(expected);
		const ids: string[] = [];
		view.state.doc.descendants((node) => {
			if (node.type === schema.nodes.paragraph) ids.push(node.attrs.id);
		});
		expect(new Set(ids).size).toBe(ids.length);
		view.destroy();
	});

	it('deletes the selected row or column and preserves remaining cells', () => {
		let fixture = setup();
		expect(run(fixture.view, 'deleteRow')).toBe(true);
		expect(grid(fixture.view)).toEqual([['r1c0', 'r1c1']]);
		fixture.view.destroy();

		fixture = setup();
		expect(run(fixture.view, 'deleteColumn')).toBe(true);
		expect(grid(fixture.view)).toEqual([['r0c1'], ['r1c1']]);
		fixture.view.destroy();
	});

	it('targets the selected nonfirst row and column', () => {
		let fixture = setup(makeTable(), 'cell-1-1');
		expect(run(fixture.view, 'deleteRow')).toBe(true);
		expect(grid(fixture.view)).toEqual([['r0c0', 'r0c1']]);
		fixture.view.destroy();

		fixture = setup(makeTable(), 'cell-1-1');
		expect(run(fixture.view, 'deleteColumn')).toBe(true);
		expect(grid(fixture.view)).toEqual([['r0c0'], ['r1c0']]);
		fixture.view.destroy();
	});

	it('replaces a final row, final column, or deleted table with a valid paragraph', () => {
		for (const [table, action] of [
			[makeTable(1, 2), 'deleteRow'],
			[makeTable(2, 1), 'deleteColumn'],
			[makeTable(1, 1), 'deleteTable'],
		] as const) {
			const { view } = setup(table);
			expect(run(view, action)).toBe(true);
			expect(view.state.doc.childCount).toBe(3);
			expect(view.state.doc.child(1).type).toBe(schema.nodes.paragraph);
			expect(view.state.doc.check()).toBeUndefined();
			view.destroy();
		}
	});

	it('records one undoable structural edit with a valid resulting selection', () => {
		const { view } = setup();
		const before = view.state.doc.toJSON();
		expect(run(view, 'rowAfter')).toBe(true);
		expect(view.state.selection.$from.parent.type).toBe(schema.nodes.paragraph);
		expect(undo(view.state, (transaction) => view.dispatch(transaction))).toBe(true);
		expect(view.state.doc.toJSON()).toEqual(before);
		view.destroy();
	});

	it('keeps consecutive structural commands in separate undo steps', () => {
		const { view } = setup();
		const initial = view.state.doc.toJSON();
		expect(run(view, 'rowAfter')).toBe(true);
		const afterFirst = view.state.doc.toJSON();
		expect(run(view, 'columnAfter')).toBe(true);
		expect(view.state.doc.toJSON()).not.toEqual(afterFirst);
		expect(undo(view.state, (transaction) => view.dispatch(transaction))).toBe(true);
		expect(view.state.doc.toJSON()).toEqual(afterFirst);
		expect(undo(view.state, (transaction) => view.dispatch(transaction))).toBe(true);
		expect(view.state.doc.toJSON()).toEqual(initial);
		view.destroy();
	});

	it('safely rejects an irregular table and readonly editing', () => {
		const irregular = schema.nodes.table.create(null, [
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, schema.nodes.paragraph.create({ id: 'cell-0-0' })),
			]),
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, schema.nodes.paragraph.create({ id: 'cell-1-0' })),
				schema.nodes.tableCell.create(null, schema.nodes.paragraph.create({ id: 'cell-1-1' })),
			]),
		]);
		const { view } = setup(irregular);
		const before = view.state.doc.toJSON();
		expect(run(view, 'columnAfter')).toBe(false);
		view.setProps({ editable: () => false });
		expect(run(view, 'deleteTable')).toBe(false);
		expect(view.state.doc.toJSON()).toEqual(before);
		view.destroy();
	});

	it('blocks row and column edits for tables marked structurally protected', () => {
		const table = makeTable();
		const protectedTable = table.type.create(
			{ ...table.attrs, structureEditable: false },
			table.content,
		);
		const { view } = setup(protectedTable);
		const before = view.state.doc.toJSON();
		expect(run(view, 'rowBefore')).toBe(false);
		expect(run(view, 'deleteColumn')).toBe(false);
		expect(view.state.doc.toJSON()).toEqual(before);
		expect(run(view, 'deleteTable')).toBe(true);
		expect(view.state.doc.child(1).type).toBe(schema.nodes.paragraph);
		view.destroy();
	});
});
