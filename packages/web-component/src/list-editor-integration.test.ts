// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel } from '@christophervr/docx-core';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
function view(editor: DocxEditorElement): EditorView {
	return (editor as unknown as { view: EditorView }).view;
}
function model(editor: DocxEditorElement): DocumentModel {
	return editor.documentModel!;
}
function click(editor: DocxEditorElement, label: string) {
	editor.shadowRoot?.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)?.click();
}

describe('list editing end to end', () => {
	afterEach(() => document.body.replaceChildren());

	it('renders a non-editable marker for a paragraph with numbering', () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [{ text: 'First item' }],
			numbering: { numId: 1, level: 0 },
		};
		model.numberingCatalog = {
			abstractNums: {
				'0': {
					id: '0',
					levels: { 0: { level: 0, start: 1, numFmt: 'decimal', lvlText: '%1.', suffix: 'tab' } },
				},
			},
			nums: { '1': { id: '1', abstractNumId: '0' } },
			warnings: [],
		};
		const editor = mount(model);
		const paragraph = editor.shadowRoot!.querySelector<HTMLElement>('.ProseMirror p')!;
		expect(paragraph.dataset.listLabel).toBe('1.\t');
		expect(paragraph.textContent).toBe('First item');
		editor.remove();
	});

	it('applies bullets to the selection via the ribbon, creating a fresh catalog entry', () => {
		const editor = mount(createDocument());
		const pmView = view(editor);
		pmView.dispatch(
			pmView.state.tr.setSelection(TextSelection.create(pmView.state.doc, 1)).scrollIntoView(),
		);
		click(editor, 'Bulleted list');
		const paragraph = model(editor).blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph');
		expect(paragraph.numbering).toBeDefined();
		expect(model(editor).numberingCatalog?.nums[String(paragraph.numbering!.numId)]).toBeDefined();
		const abstractId =
			model(editor).numberingCatalog!.nums[String(paragraph.numbering!.numId)].abstractNumId;
		expect(model(editor).numberingCatalog!.abstractNums[abstractId].levels[0].numFmt).toBe(
			'bullet',
		);
		click(editor, 'Bulleted list');
		const toggledOff = model(editor).blocks[0];
		if (toggledOff.type !== 'paragraph') throw new Error('expected a paragraph');
		expect(toggledOff.numbering).toBeUndefined();
		editor.remove();
	});

	it('increases and decreases the outline level from the ribbon', () => {
		const editor = mount(createDocument());
		const pmView = view(editor);
		pmView.dispatch(pmView.state.tr.setSelection(TextSelection.create(pmView.state.doc, 1)));
		click(editor, 'Numbered list');
		click(editor, 'Increase list level');
		let paragraph = model(editor).blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph');
		expect(paragraph.numbering?.level).toBe(1);
		click(editor, 'Decrease list level');
		paragraph = model(editor).blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph');
		expect(paragraph.numbering?.level).toBe(0);
		click(editor, 'Remove list');
		paragraph = model(editor).blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected a paragraph');
		expect(paragraph.numbering).toBeUndefined();
		editor.remove();
	});
});
