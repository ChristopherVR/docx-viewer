// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, type Table } from 'docx-core';
import { history, redo, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createBordersDialog } from './borders-dialog';
import { modelToDoc, docToModel } from './model-adapter';
import { applyCellBorderSettings, tableBorderContext } from './table-border-commands';
import { readCellFill } from './table-shading';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
function setup(style = false) {
	const model = createDocument();
	const cell = (id: string) => ({
		paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: id }] }],
	});
	model.blocks = [
		{
			type: 'table',
			id: 'table',
			rows: [
				[cell('a'), cell('b')],
				[cell('c'), cell('d')],
			],
			...(style ? { style: 'Shaded' } : {}),
		},
	];
	if (style)
		model.tableStyles = {
			warnings: [],
			styles: { Shaded: { id: 'Shaded', conditional: {}, shadingFill: '#abcdef' } },
		};
	const doc = modelToDoc(model);
	let caret = 0;
	doc.descendants((node, pos) => {
		if (node.attrs.id === 'a') caret = pos + 1;
	});
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, caret),
			plugins: [history()],
		}),
	});
	views.push(view);
	return { view, model };
}
const settings = {
	scope: 'cell' as const,
	sides: { top: true },
	pen: { style: 'single', sizeEighthPoints: 4 },
	fill: '#123456',
};

describe('editable table shading', () => {
	it('applies fill and borders to the current cell with one undo and redo', () => {
		const { view, model } = setup();
		const original = view.state.doc;
		expect(applyCellBorderSettings(view, settings)).toBe(true);
		let table = docToModel(view.state.doc, model).blocks[0] as Table;
		expect(table.rows[0]![0]!.shadingFill).toBe('#123456');
		expect(table.rows[0]![0]!.borders?.top?.style).toBe('single');
		expect(table.rows[0]![1]!.shadingFill).toBeUndefined();
		undo(view.state, (tr) => view.dispatch(tr));
		expect(view.state.doc.eq(original)).toBe(true);
		redo(view.state, (tr) => view.dispatch(tr));
		table = docToModel(view.state.doc, model).blocks[0] as Table;
		expect(table.rows[0]![0]!.shadingFill).toBe('#123456');
	});

	it('keeps inherited fills out of direct formatting until explicitly edited, and No Color overrides a style', () => {
		const { view, model } = setup(true);
		expect(readCellFill(tableBorderContext(view.state)!, view.state, 'cell')).toBe('#abcdef');
		expect(
			(docToModel(view.state.doc, model).blocks[0] as Table).rows[0]![0]!.shadingFill,
		).toBeUndefined();
		applyCellBorderSettings(view, { ...settings, fill: null });
		const table = docToModel(view.state.doc, model).blocks[0] as Table;
		expect(table.rows[0]![0]!.shadingFill).toBe('auto');
		expect(table.rows[0]![1]!.shadingFill).toBeUndefined();
		expect(
			modelToDoc({ ...model, blocks: [table] }).firstChild!.firstChild!.firstChild!.attrs
				.shadingFill,
		).toBe('auto');
	});

	it('applies Table scope to every cell and leaves fill unchanged when omitted', () => {
		const { view, model } = setup();
		applyCellBorderSettings(view, { ...settings, scope: 'table' });
		applyCellBorderSettings(view, { scope: 'cell', sides: {}, pen: settings.pen });
		const table = docToModel(view.state.doc, model).blocks[0] as Table;
		expect(table.rows.flat().map((cell) => cell.shadingFill)).toEqual(Array(4).fill('#123456'));
	});

	it('shows mixed table fills without changing them if OK is pressed, then applies an explicit fill', () => {
		const { view, model } = setup();
		applyCellBorderSettings(view, settings);
		const dialog = createBordersDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const scope = dialog.element.querySelector<HTMLSelectElement>('[aria-label="Apply to"]')!;
		scope.value = 'table';
		scope.dispatchEvent(new Event('change'));
		const noColor = dialog.element.querySelector<HTMLInputElement>('[aria-label="No Color"]')!;
		expect(noColor.indeterminate).toBe(true);
		const ok = () =>
			[...dialog.element.querySelectorAll('button')]
				.find((button) => button.textContent === 'OK')!
				.click();
		ok();
		expect(
			(docToModel(view.state.doc, model).blocks[0] as Table).rows[0]![1]!.shadingFill,
		).toBeUndefined();
		dialog.open();
		const fill = dialog.element.querySelector<HTMLInputElement>('[aria-label="Fill"]')!;
		expect(fill.disabled).toBe(false);
		fill.value = '#ff0000';
		fill.dispatchEvent(new Event('input'));
		ok();
		expect((docToModel(view.state.doc, model).blocks[0] as Table).rows[0]![0]!.shadingFill).toBe(
			'#ff0000',
		);
	});

	it('rejects shading changes in viewing mode and protected tables', () => {
		const { view } = setup();
		view.setProps({ editable: () => false });
		expect(applyCellBorderSettings(view, settings)).toBe(false);
		view.setProps({ editable: () => true });
		view.dispatch(
			view.state.tr.setNodeMarkup(0, undefined, {
				...view.state.doc.firstChild!.attrs,
				structureEditable: false,
			}),
		);
		expect(applyCellBorderSettings(view, settings)).toBe(false);
	});
});
