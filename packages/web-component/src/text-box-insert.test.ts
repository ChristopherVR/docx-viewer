// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { undo } from 'prosemirror-history';
import { history } from 'prosemirror-history';
import { createDocument, type Paragraph } from 'docx-core';
import { EditorState, NodeSelection, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import { createTextBoxDialog } from './text-box-dialog';
import { insertTextBox, selectedTextBox, updateTextBox } from './text-box-commands';

function editor() {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Hello' }] }];
	const doc = modelToDoc(model);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, 3),
			plugins: [history()],
		}),
	});
	return { model, view };
}
const settings = { lines: ['One', 'Two'], widthPx: 192, heightPx: 96, border: true };

describe('text box commands', () => {
	it('inserts an editable box that converts to a text-box image run', () => {
		const { model, view } = editor();
		expect(insertTextBox(view, settings)).toBe(true);
		const next = docToModel(view.state.doc, model);
		const image = (next.blocks[0] as Paragraph).runs.find((run) => run.image)!.image!;
		expect(image).toMatchObject({
			unsupported: 'Text box',
			textBoxText: ['One', 'Two'],
			textBoxEditable: true,
			textBoxBorder: true,
			widthPx: 192,
			heightPx: 96,
		});
		expect((next.blocks[0] as Paragraph).runs.map((run) => run.text).join('')).toBe('Hello');
		undo(view.state, view.dispatch);
		expect(view.state.doc.textContent).toBe('Hello');
		expect(JSON.stringify(view.state.doc.toJSON())).not.toContain('Text box');
	});

	it('keeps text typed after an inline image in its own run', () => {
		const { model, view } = editor();
		insertTextBox(view, settings);
		const runs = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs;
		expect(runs.map((run) => [run.text, Boolean(run.image)])).toEqual([
			['He', false],
			['', true],
			['llo', false],
		]);
	});

	it('edits the selected box only, as one undo step', () => {
		const { model, view } = editor();
		insertTextBox(view, settings);
		let pos = -1;
		view.state.doc.descendants((node, at) => {
			if (node.type.name === 'image') pos = at;
		});
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
		expect(selectedTextBox(view.state)!.settings.lines).toEqual(['One', 'Two']);
		expect(updateTextBox(view, pos, { ...settings, lines: ['Only'], border: false })).toBe(true);
		const image = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs.find(
			(run) => run.image,
		)!.image!;
		expect(image).toMatchObject({ textBoxText: ['Only'], textBoxBorder: false });
		undo(view.state, view.dispatch);
		expect(selectedTextBox(view.state)?.settings.lines ?? ['gone']).toEqual(['One', 'Two']);
		view.setProps({ editable: () => false });
		expect(insertTextBox(view, settings)).toBe(false);
	});
});

describe('Text box dialog', () => {
	it('inserts from the dialog fields and validates the size', () => {
		const { model, view } = editor();
		const dialog = createTextBoxDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const field = (name: string) =>
			dialog.element.querySelector<HTMLInputElement & HTMLTextAreaElement>(
				`[aria-label="${name}"]`,
			)!;
		const ok = [...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!;
		field('Text').value = 'First\nSecond';
		field('Width (inches)').value = '30';
		field('Width (inches)').dispatchEvent(new Event('input'));
		expect(ok.disabled).toBe(true);
		field('Width (inches)').value = '3';
		field('Width (inches)').dispatchEvent(new Event('input'));
		field('Outline').checked = false;
		ok.click();
		const image = (docToModel(view.state.doc, model).blocks[0] as Paragraph).runs.find(
			(run) => run.image,
		)!.image!;
		expect(image).toMatchObject({
			textBoxText: ['First', 'Second'],
			widthPx: 288,
			heightPx: 96,
			textBoxBorder: false,
		});
		expect(dialog.isOpen).toBe(false);
		dialog.element.remove();
	});
});
