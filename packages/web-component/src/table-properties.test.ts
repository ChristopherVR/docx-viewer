// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, twips, signedTwips, type Table } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo } from 'prosemirror-history';
import { modelToDoc, docToModel } from './model-adapter';
import { applyTableProperties } from './table-properties';
import { executeTableCommand } from './table-commands';
import { createTablePropertiesDialog } from './table-properties-dialog';

const views: EditorView[] = [];
afterEach(() => {
	views.forEach((view) => view.destroy());
	views.length = 0;
	document.body.replaceChildren();
});
function setup(allRows = false, complex = false) {
	const model = createDocument();
	const table: Table = {
		type: 'table',
		id: 'table',
		structureEditable: !complex,
		cellMargins: { left: signedTwips(108) },
		rowProperties: [
			{ heightTwips: twips(360), heightRule: 'exact', header: true },
			{ heightTwips: twips(720), heightRule: 'atLeast', cantSplit: true },
		],
		rows: [0, 1].map((index) => [
			{
				paragraphs: [{ type: 'paragraph', id: `p${index}`, runs: [{ text: `row${index}` }] }],
				margins: { top: signedTwips(40) },
			},
		]),
	};
	model.blocks = [table];
	const doc = modelToDoc(model);
	let last = 0;
	doc.descendants((node, pos) => {
		if (node.attrs.id === 'p1') last = pos + 1;
	});
	const view = new EditorView(document.body.appendChild(document.createElement('div')), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 4, allRows ? last + 1 : 4),
			plugins: [history()],
		}),
	});
	views.push(view);
	const read = () => docToModel(view.state.doc, model).blocks[0] as Table;
	const dialog = createTablePropertiesDialog(() => view);
	document.body.append(dialog.element);
	const input = (name: string) =>
		dialog.element.querySelector<HTMLInputElement>(`[aria-label="${name}"]`)!;
	const change = (name: string, value: string | boolean) => {
		const control = input(name);
		if (typeof value === 'boolean') control.checked = value;
		else control.value = value;
		control.dispatchEvent(new Event('input', { bubbles: true }));
	};
	const submit = () =>
		[...dialog.element.querySelectorAll('button')]
			.find((button) => button.textContent === 'OK')!
			.click();
	return { view, model, read, dialog, input, change, submit };
}
describe('Table Properties', () => {
	it('saves edited properties into a new DOCX and patches an opened package on the next save', async () => {
		const { view, model } = setup();
		applyTableProperties(
			view,
			{ header: true, cantSplit: true, heightTwips: twips(540), heightRule: 'exact' },
			{ bottom: signedTwips(144) },
		);
		const loaded = await loadDocx(await saveDocx(docToModel(view.state.doc, model)));
		expect((loaded.model.blocks[0] as Table).rowProperties![0]).toEqual({
			header: true,
			cantSplit: true,
			heightTwips: 540,
			heightRule: 'exact',
		});
		const state = EditorState.create({ doc: modelToDoc(loaded.model) });
		const updated = state.tr.setNodeMarkup(1, undefined, {
			...state.doc.nodeAt(1)!.attrs,
			properties: { header: false },
		});
		const reopened = await loadDocx(await loaded.save(docToModel(updated.doc, loaded.model)));
		expect((reopened.model.blocks[0] as Table).rowProperties![0]).toEqual({});
		expect((reopened.model.blocks[0] as Table).cellMargins).toEqual({ left: 108, bottom: 144 });
	});
	it('applies row and default margin changes in one undo step and keeps cell overrides', () => {
		const { view, read } = setup();
		const before = read();
		expect(
			applyTableProperties(
				view,
				{ heightTwips: twips(500), cantSplit: true },
				{ top: signedTwips(144) },
			),
		).toBe(true);
		expect(read().rowProperties).toEqual([
			{ heightTwips: 500, heightRule: 'exact', header: true, cantSplit: true },
			before.rowProperties![1],
		]);
		expect(read().cellMargins).toEqual({ left: 108, top: 144 });
		expect(read().rows[0]![0]!.margins).toEqual({ top: 40 });
		expect(undo(view.state, view.dispatch)).toBe(true);
		expect(read()).toEqual(before);
	});
	it('keeps properties attached to surviving rows through insertion, deletion and column edits', () => {
		const { view, read } = setup();
		executeTableCommand(view, 'rowBefore');
		expect(read().rowProperties).toEqual([
			{},
			{ heightTwips: 360, heightRule: 'exact', header: true },
			{ heightTwips: 720, heightRule: 'atLeast', cantSplit: true },
		]);
		executeTableCommand(view, 'columnAfter');
		expect(read().rowProperties![2]).toMatchObject({ cantSplit: true, heightTwips: 720 });
		executeTableCommand(view, 'deleteRow');
		expect(read().rowProperties![0]).toMatchObject({ header: true, heightRule: 'exact' });
	});
	it('shows mixed rows and changes only touched fields', () => {
		const { dialog, input, change, submit, read } = setup(true);
		dialog.open();
		expect(input('Height (inches)').value).toBe('');
		expect(input('Allow row to break across pages').indeterminate).toBe(true);
		expect(input('Repeat as header row at the top of each page').indeterminate).toBe(true);
		change('Allow row to break across pages', false);
		submit();
		expect(read().rowProperties).toEqual([
			{ heightTwips: 360, heightRule: 'exact', header: true, cantSplit: true },
			{ heightTwips: 720, heightRule: 'atLeast', cantSplit: true },
		]);
	});
	it('clears specified height, validates measurements and cancels without editing', () => {
		const { dialog, change, submit, read } = setup();
		dialog.open();
		change('Top', '-1');
		submit();
		expect(dialog.isOpen).toBe(true);
		expect(read().cellMargins).toEqual({ left: 108 });
		dialog.close();
		dialog.open();
		change('Specify height', false);
		submit();
		expect(read().rowProperties![0]).toEqual({ header: true });
	});
	it('refuses formatting in complex tables and read-only views', () => {
		const { view, dialog } = setup(false, true);
		dialog.open();
		expect(dialog.isOpen).toBe(false);
		expect(applyTableProperties(view, { header: true }, {})).toBe(false);
		const simple = setup();
		simple.view.setProps({ editable: () => false });
		expect(applyTableProperties(simple.view, { header: true }, {})).toBe(false);
	});
});
