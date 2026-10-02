// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel } from 'docx-core';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { must, paragraphAt } from './test-support';

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
		expect(editor.shadowRoot!.querySelector<HTMLElement>('.ProseMirror p')!.dataset.listLabel).toBe(
			'•\t',
		);
		const paragraph = paragraphAt(model(editor).blocks, 0);
		expect(paragraph.numbering).toBeDefined();
		const numId = String(must(paragraph.numbering).numId);
		const catalog = must(model(editor).numberingCatalog, 'numbering catalog');
		expect(catalog.nums[numId]).toBeDefined();
		const abstractId = must(catalog.nums[numId]).abstractNumId;
		const abstractNum = must(catalog.abstractNums[abstractId], 'abstract numbering');
		expect(must(abstractNum.levels[0], 'level 0').numFmt).toBe('bullet');
		click(editor, 'Bulleted list');
		const toggledOff = paragraphAt(model(editor).blocks, 0);
		expect(toggledOff.numbering).toBeUndefined();
		expect(
			editor.shadowRoot!.querySelector<HTMLElement>('.ProseMirror p')!.dataset.listLabel ?? '',
		).toBe('');
		editor.remove();
	});

	it('increases and decreases the outline level from the ribbon', () => {
		const editor = mount(createDocument());
		const pmView = view(editor);
		pmView.dispatch(pmView.state.tr.setSelection(TextSelection.create(pmView.state.doc, 1)));
		click(editor, 'Numbered list');
		click(editor, 'Increase list level');
		expect(editor.shadowRoot!.querySelector<HTMLElement>('.ProseMirror p')!.dataset.listLabel).toBe(
			'a.\t',
		);
		let paragraph = paragraphAt(model(editor).blocks, 0);
		expect(paragraph.numbering?.level).toBe(1);
		click(editor, 'Decrease list level');
		paragraph = paragraphAt(model(editor).blocks, 0);
		expect(paragraph.numbering?.level).toBe(0);
		click(editor, 'Remove list');
		paragraph = paragraphAt(model(editor).blocks, 0);
		expect(paragraph.numbering).toBeUndefined();
		editor.remove();
	});
});
