// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph } from 'docx-core';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { must } from './test-support';

registerDocxEditor();

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const view = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const firstParagraph = (editor: DocxEditorElement) => editor.documentModel!.blocks[0] as Paragraph;
const click = (editor: DocxEditorElement, label: string) =>
	editor.shadowRoot?.querySelector<HTMLButtonElement>(`[aria-label="${label}"]`)?.click();

function withText(text: string, trackChanges = false): DocumentModel {
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text }] };
	model.trackChanges = trackChanges;
	return model;
}

describe('review editing end to end', () => {
	afterEach(() => document.body.replaceChildren());

	it('records typing as a tracked insertion and accepts it', () => {
		const editor = mount(withText('Hello', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.insertText(' world', 6));
		expect(pm.state.doc.textContent).toBe('Hello world');
		const runs = firstParagraph(editor).runs;
		expect(runs.map((run) => run.text).join('')).toBe('Hello world');
		expect(runs.find((run) => run.text === ' world')?.revision?.kind).toBe('insert');
		click(editor, 'Accept all');
		const accepted = firstParagraph(editor).runs;
		expect(accepted.map((run) => run.text).join('')).toBe('Hello world');
		expect(accepted.some((run) => run.revision)).toBe(false);
	});

	it('keeps deleted text visible as a tracked deletion and rejects it', () => {
		const editor = mount(withText('Hello world', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.delete(6, 12));
		expect(pm.state.doc.textContent).toBe('Hello world');
		const deleted = firstParagraph(editor).runs.find((run) => run.revision?.kind === 'delete');
		expect(deleted?.text).toBe(' world');
		click(editor, 'Reject all');
		const rejected = firstParagraph(editor).runs;
		expect(rejected.map((run) => run.text).join('')).toBe('Hello world');
		expect(rejected.some((run) => run.revision)).toBe(false);
	});

	it('removes the current author’s own pending insertion outright', () => {
		const editor = mount(withText('Hi', true));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.insertText('!', 3));
		pm.dispatch(pm.state.tr.delete(3, 4));
		expect(pm.state.doc.textContent).toBe('Hi');
		expect(firstParagraph(editor).runs.some((run) => run.revision)).toBe(false);
	});

	it('toggles track changes from the Review ribbon', () => {
		const editor = mount(withText('Text'));
		click(editor, 'Track changes');
		expect(editor.documentModel!.trackChanges).toBe(true);
		click(editor, 'Track changes');
		expect(editor.documentModel!.trackChanges).toBe(false);
	});

	it('adds a comment anchored to the selection from the comments pane', () => {
		const editor = mount(withText('Review me'));
		const pm = view(editor);
		pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, 1, 7)));
		click(editor, 'Add comment');
		const input = editor.shadowRoot!.querySelector<HTMLTextAreaElement>(
			'[aria-label="New comment"]',
		)!;
		expect(input).toBeTruthy();
		input.value = 'Please check';
		input.dispatchEvent(new Event('input'));
		const submit = [...editor.shadowRoot!.querySelectorAll<HTMLButtonElement>('button')].find(
			(button) => button.textContent === 'Add comment' && button.closest('[aria-label="Comments"]'),
		);
		submit?.click();
		const comments = editor.documentModel!.comments ?? [];
		expect(comments).toHaveLength(1);
		const [comment] = comments;
		expect(must(comment).text).toBe('Please check');
		const anchored = firstParagraph(editor).runs.find((run) =>
			run.commentIds?.includes(must(comment).id),
		);
		expect(anchored?.text).toBe('Review');
	});
});
