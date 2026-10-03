// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, signedTwips, type Table } from 'docx-core';
import { history, redo, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createTablePropertiesDialog } from './table-properties-dialog';
import { applyTableProperties } from './table-properties';
import { modelToDoc, docToModel } from './model-adapter';
import { tableCellStyle, tableStyle } from './table-render';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
function setup(multiple = false) {
	const model = createDocument();
	const cell = (id: string, margins?: Table['rows'][number][number]['margins']) => ({
		paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: id }] }],
		...(margins ? { margins } : {}),
	});
	const table: Table = {
		type: 'table',
		id: 't',
		cellMargins: { left: signedTwips(216), top: signedTwips(72) },
		rows: [
			[
				cell('a', { top: signedTwips(144) }),
				cell('b', { top: signedTwips(0), right: signedTwips(108) }),
			],
			[cell('c'), cell('d', { bottom: signedTwips(54) })],
		],
	};
	model.blocks = [table];
	const doc = modelToDoc(model);
	const positions = new Map<string, number>();
	doc.descendants((node, pos) => {
		if (node.type.name === 'paragraph') positions.set(String(node.attrs.id), pos + 1);
	});
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(
				doc,
				positions.get('a')!,
				positions.get(multiple ? 'c' : 'a')!,
			),
			plugins: [history()],
		}),
	});
	views.push(view);
	const read = () => docToModel(view.state.doc, model).blocks[0] as Table;
	const dialog = createTablePropertiesDialog(() => view);
	document.body.append(dialog.element);
	const control = (label: string) =>
		dialog.element.querySelector<HTMLInputElement>(`[aria-label="${label}"]`)!;
	const submit = () =>
		[...dialog.element.querySelectorAll('button')]
			.find((button) => button.textContent === 'OK')!
			.click();
	return { view, model, read, dialog, control, submit };
}

describe('selected cell margins', () => {
	it('patches selected cells only, preserving untouched overrides and combining default/row/cell changes in one undo', () => {
		const { view, read } = setup(true);
		const original = view.state.doc;
		applyTableProperties(
			view,
			{ cantSplit: true },
			{ left: signedTwips(288) },
			{ right: signedTwips(432) },
		);
		const table = read();
		expect(table.cellMargins).toEqual({ left: 288, top: 72 });
		expect(table.rows.flat().map((cell) => cell.margins)).toEqual([
			{ top: 144, right: 432 },
			{ top: 0, right: 432 },
			{ right: 432 },
			{ bottom: 54 },
		]);
		expect(table.rowProperties).toEqual([{ cantSplit: true }, { cantSplit: true }]);
		expect(undo(view.state, view.dispatch)).toBe(true);
		expect(view.state.doc.eq(original)).toBe(true);
		expect(redo(view.state, view.dispatch)).toBe(true);
		expect(read().rows[0]![0]!.margins).toEqual({ top: 144, right: 432 });
	});

	it('removes cell overrides for inheritance and writes explicit zero independently', () => {
		const { view, read } = setup();
		applyTableProperties(view, {}, {}, { top: signedTwips(0) });
		expect(read().rows[0]![0]!.margins).toEqual({ top: 0 });
		applyTableProperties(view, {}, {}, null);
		expect(read().rows[0]![0]!.margins).toBeUndefined();
		expect(read().rows[0]![1]!.margins).toEqual({ top: 0, right: 108 });
		undo(view.state, view.dispatch);
		expect(read().rows[0]![0]!.margins).toEqual({ top: 0 });
	});

	it('shows mixed effective margins and inheritance without flattening untouched fields', () => {
		const { dialog, control, submit, view, read } = setup(true);
		const original = view.state.doc;
		dialog.open();
		expect(control('Cell top').value).toBe('');
		expect(control('Cell left').value).toBe('0.15');
		expect(control('Use table defaults').indeterminate).toBe(true);
		submit();
		expect(view.state.doc.eq(original)).toBe(true);
		dialog.open();
		const right = control('Cell right');
		right.value = '0.25';
		right.dispatchEvent(new Event('input'));
		submit();
		expect(
			read()
				.rows.flat()
				.map((cell) => cell.margins),
		).toEqual([{ top: 144, right: 360 }, { top: 0, right: 360 }, { right: 360 }, { bottom: 54 }]);
	});

	it('Use table defaults resets the selected cells and validation prevents invalid edits', () => {
		const { dialog, control, submit, read } = setup();
		dialog.open();
		const left = control('Cell left');
		left.value = '-0.1';
		left.dispatchEvent(new Event('input'));
		submit();
		expect(dialog.isOpen).toBe(true);
		expect(read().rows[0]![0]!.margins).toEqual({ top: 144 });
		const defaults = control('Use table defaults');
		defaults.checked = true;
		defaults.dispatchEvent(new Event('change'));
		expect(left.disabled).toBe(true);
		submit();
		expect(dialog.isOpen).toBe(false);
		expect(read().rows[0]![0]!.margins).toBeUndefined();
	});

	it('renders partial cell overrides through per-side table defaults and keeps inherited values out of the model', () => {
		const { view, read } = setup();
		const attrs = view.state.doc.firstChild!.attrs;
		expect(tableStyle(attrs)).toContain('--dve-cell-padding-left:14.4px');
		expect(tableStyle(attrs)).toContain('--dve-cell-padding-top:4.8px');
		const cellAttrs = view.state.doc.firstChild!.firstChild!.firstChild!.attrs;
		expect(tableCellStyle(cellAttrs)).toContain(
			'padding:9.6px var(--dve-cell-padding-right, 7.2px)',
		);
		expect(tableCellStyle(cellAttrs)).toContain('var(--dve-cell-padding-left, 7.2px)');
		expect(read().rows[0]![0]!.margins).toEqual({ top: 144 });
	});
});
