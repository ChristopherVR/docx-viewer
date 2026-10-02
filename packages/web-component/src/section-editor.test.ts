// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, type DocumentModel } from 'docx-core';
import { TextSelection } from 'prosemirror-state';
import { undo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { at, must } from './test-support';

registerDocxEditor();

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const view = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
function choose(editor: DocxEditorElement, label: string, value: string) {
	const select = editor.shadowRoot!.querySelector<HTMLSelectElement>(`[aria-label="${label}"]`)!;
	select.value = value;
	select.dispatchEvent(new Event('change'));
}
function twoParagraphs(): DocumentModel {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'p1', runs: [{ text: 'Portrait part' }] },
		{ type: 'paragraph', id: 'p2', runs: [{ text: 'Landscape part' }] },
	];
	return model;
}

describe('section editing in the editor', () => {
	afterEach(() => document.body.replaceChildren());

	it('makes page setup undoable, including the first change to a new document', () => {
		const editor = mount(createDocument());
		choose(editor, 'Orientation', 'landscape');
		expect(editor.documentModel!.sections?.[0]?.orientation).toBe('landscape');
		expect(editor.documentModel!.page.width).toBe(1056);
		undo(view(editor).state, view(editor).dispatch);
		expect(editor.documentModel!.sections).toBeUndefined();
		expect(editor.documentModel!.page.width).toBe(816);
	});

	it('inserts a section break and gives the second section its own orientation', async () => {
		const editor = mount(twoParagraphs());
		const pm = view(editor);
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 3)));
		editor
			.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Next page section break"]')!
			.click();
		expect(editor.documentModel!.sections?.map((section) => section.endsAtBlockId)).toEqual([
			'p1',
			'p2',
		]);
		expect(
			editor.shadowRoot!.querySelector('[data-section-break]')?.getAttribute('data-section-break'),
		).toBe('Section Break (Next Page)');
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 18)));
		choose(editor, 'Orientation', 'landscape');
		const sections = must(editor.documentModel!.sections, 'sections');
		const [first, second] = [at(sections, 0), at(sections, 1)];
		expect([first.orientation, second.orientation]).toEqual(['portrait', 'landscape']);
		const zip = await JSZip.loadAsync(await editor.saveBytes());
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toMatch(/Portrait part.*<\/w:p>.*w:orient="landscape"/);
		expect(xml).toMatch(/<w:pPr><w:sectPr>/);
	});

	it('merges sections when the paragraph ending a section is deleted', () => {
		const editor = mount(twoParagraphs());
		const pm = view(editor);
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 3)));
		editor
			.shadowRoot!.querySelector<HTMLButtonElement>('[aria-label="Continuous section break"]')!
			.click();
		expect(editor.documentModel!.sections).toHaveLength(2);
		pm.dispatch(pm.state.tr.delete(0, pm.state.doc.child(0).nodeSize));
		expect(editor.documentModel!.sections).toHaveLength(1);
	});

	it('sets a different first page and restarted Roman page numbers for the current section', () => {
		const editor = mount(twoParagraphs());
		const toggle = editor.shadowRoot!.querySelector<HTMLButtonElement>(
			'[aria-label="Different first page"]',
		)!;
		toggle.click();
		expect(editor.documentModel!.sections?.[0]?.titlePage).toBe(true);
		expect(toggle.getAttribute('aria-pressed')).toBe('true');
		choose(editor, 'Page number format', 'lowerRoman');
		choose(editor, 'Page numbering', 'restart');
		expect(editor.documentModel!.sections?.[0]?.pageNumbering).toEqual({
			format: 'lowerRoman',
			start: 1,
		});
		choose(editor, 'Page number format', 'decimal');
		expect(editor.documentModel!.sections?.[0]?.pageNumbering).toEqual({ start: 1 });
		undo(view(editor).state, view(editor).dispatch);
		expect(editor.documentModel!.sections?.[0]?.pageNumbering?.format).toBe('lowerRoman');
	});

	it('sets vertical alignment and the document-wide odd and even pages setting', async () => {
		const editor = mount(twoParagraphs());
		choose(editor, 'Vertical alignment', 'center');
		expect(editor.documentModel!.sections?.[0]?.verticalAlign).toBe('center');
		const toggle = editor.shadowRoot!.querySelector<HTMLButtonElement>(
			'[aria-label="Different odd and even pages"]',
		)!;
		toggle.click();
		expect(editor.documentModel!.evenAndOddHeaders).toBe(true);
		expect(toggle.getAttribute('aria-pressed')).toBe('true');
		const zip = await JSZip.loadAsync(await editor.saveBytes());
		expect(await zip.file('word/settings.xml')!.async('string')).toContain(
			'<w:evenAndOddHeaders/>',
		);
		expect(await zip.file('word/document.xml')!.async('string')).toContain(
			'<w:vAlign w:val="center"/>',
		);
		undo(view(editor).state, view(editor).dispatch);
		expect(editor.documentModel!.evenAndOddHeaders).toBeUndefined();
	});
});
