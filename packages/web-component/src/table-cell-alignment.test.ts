// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, signedTwips, type Table, type TableCell } from 'docx-core';
import { history, redo, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { createTablePropertiesDialog } from './table-properties-dialog';
import { applyTableProperties } from './table-properties';
import { modelToDoc, docToModel } from './model-adapter';

const views: EditorView[] = [];
afterEach(() => {
	for (const view of views.splice(0)) view.destroy();
	document.body.replaceChildren();
});
function setup(multiple = false, imported: TableCell['verticalAlign'] = 'bottom') {
	const model = createDocument();
	const cell = (id: string, verticalAlign?: TableCell['verticalAlign']) => ({
		paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: id }] }],
		...(verticalAlign ? { verticalAlign } : {}),
		margins: { left: signedTwips(144) },
	});
	model.blocks = [
		{
			type: 'table',
			id: 't',
			rows: [
				[cell('a', imported), cell('b', 'center')],
				[cell('c'), cell('d', 'bottom')],
			],
		},
	];
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
	const alignment = () =>
		dialog.element.querySelector<HTMLSelectElement>('[aria-label="Cell vertical alignment"]')!;
	const choose = (value: string) => {
		alignment().value = value;
		alignment().dispatchEvent(new Event('change'));
	};
	const submit = () =>
		[...dialog.element.querySelectorAll('button')]
			.find((button) => button.textContent === 'OK')!
			.click();
	return { view, read, dialog, alignment, choose, submit };
}

describe('selected cell vertical alignment', () => {
	it('applies to selected cells and joins row and margin changes into one undo/redo', () => {
		const { view, read } = setup(true);
		const original = view.state.doc;
		applyTableProperties(view, { cantSplit: true }, {}, { top: signedTwips(72) }, 'center');
		expect(
			read()
				.rows.flat()
				.map((cell) => cell.verticalAlign),
		).toEqual(['center', 'center', 'center', 'bottom']);
		expect(read().rows[0]![0]!.margins).toEqual({ left: 144, top: 72 });
		expect(read().rowProperties).toEqual([{ cantSplit: true }, { cantSplit: true }]);
		undo(view.state, view.dispatch);
		expect(view.state.doc.eq(original)).toBe(true);
		redo(view.state, view.dispatch);
		expect(read().rows[0]![0]!.verticalAlign).toBe('center');
	});

	it('writes explicit Top and removes an override for Default alignment', () => {
		const { view, read, dialog, alignment, choose, submit } = setup();
		dialog.open();
		expect(alignment().value).toBe('bottom');
		choose('top');
		submit();
		expect(read().rows[0]![0]!.verticalAlign).toBe('top');
		dialog.open();
		choose('default');
		submit();
		expect(read().rows[0]![0]!.verticalAlign).toBeUndefined();
		expect(read().rows[0]![1]!.verticalAlign).toBe('center');
		undo(view.state, view.dispatch);
		expect(read().rows[0]![0]!.verticalAlign).toBe('top');
	});

	it('leaves mixed alignments unchanged until the control is changed', () => {
		const { view, read, dialog, alignment, choose, submit } = setup(true);
		const original = view.state.doc;
		dialog.open();
		expect(alignment().value).toBe('');
		submit();
		expect(view.state.doc.eq(original)).toBe(true);
		dialog.open();
		choose('bottom');
		submit();
		expect(
			read()
				.rows.flat()
				.map((cell) => cell.verticalAlign),
		).toEqual(['bottom', 'bottom', 'bottom', 'bottom']);
	});

	it('shows imported Justified truthfully as a preservation-only option and leaves it intact', () => {
		const { read, dialog, alignment, submit } = setup(false, 'both');
		dialog.open();
		expect(alignment().value).toBe('both');
		const option = alignment().querySelector<HTMLOptionElement>('[value="both"]')!;
		expect(option.textContent).toBe('Justified (imported)');
		expect(option.disabled).toBe(true);
		submit();
		expect(read().rows[0]![0]!.verticalAlign).toBe('both');
	});

	it('does not apply changes in viewing mode or a protected table', () => {
		const { view } = setup();
		view.setProps({ editable: () => false });
		expect(applyTableProperties(view, {}, {}, undefined, 'center')).toBe(false);
		view.setProps({ editable: () => true });
		view.dispatch(
			view.state.tr.setNodeMarkup(0, undefined, {
				...view.state.doc.firstChild!.attrs,
				structureEditable: false,
			}),
		);
		expect(applyTableProperties(view, {}, {}, undefined, 'center')).toBe(false);
	});
});
