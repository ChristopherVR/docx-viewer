// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from 'docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text: 'Hello chrome' }] };
	model.warnings = ['Pictures are shown as placeholders.'];
	editor.documentModel = model;
	editor.fileName = 'Report.docx';
	document.body.append(editor);
	return editor;
}
const root = (editor: DocxEditorElement) => editor.shadowRoot!;

describe('Word window chrome', () => {
	it('File Options changes live view and spelling preferences without editing the document', async () => {
		const editor = mount();
		const r = root(editor);
		r.querySelector<HTMLButtonElement>('.dve-file-tab')!.click();
		[...r.querySelectorAll<HTMLButtonElement>('.dve-backstage-nav-item')]
			.find((button) => button.textContent === 'Options')!
			.click();
		const toggle = async (label: string) => {
			r.querySelector<HTMLInputElement>(
				`.dve-backstage-content input[aria-label="${label}"]`,
			)!.click();
			await Promise.resolve();
		};
		await toggle('Ruler');
		expect(r.querySelector('.dve-ruler')).not.toBeNull();
		await toggle('Show paragraph marks');
		expect(r.querySelector('.dve-paper')?.hasAttribute('data-show-marks')).toBe(true);
		await toggle('Page thumbnails');
		expect(editor.hasAttribute('show-thumbnails')).toBe(true);
		await toggle('Spelling');
		expect(
			(editor as unknown as { view: import('prosemirror-view').EditorView }).view.dom.spellcheck,
		).toBe(false);
		expect(editor.dirty).toBe(false);
		await toggle('Ruler');
		expect(r.querySelector('.dve-ruler')).toBeNull();
	});
	afterEach(() => document.body.replaceChildren());

	it('shows the file name and tracks unsaved changes in the title bar', () => {
		const editor = mount();
		expect(root(editor).querySelector('.dve-filename')?.textContent).toBe('Report.docx');
		expect(root(editor).querySelector('.dve-save-state')?.textContent).toBe('Saved');
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('!', 13));
		expect(root(editor).querySelector('.dve-save-state')?.textContent).toBe('Unsaved changes');
	});

	it('runs a ribbon command found through "Tell me what you want to do"', () => {
		const editor = mount();
		const input = root(editor).querySelector<HTMLInputElement>('.dve-tellme input')!;
		input.value = 'bulleted';
		input.dispatchEvent(new Event('input'));
		const options = [...root(editor).querySelectorAll('.dve-tellme-results [role=option]')];
		expect(options.map((option) => option.textContent)).toEqual(['Bulleted list']);
		input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(editor.documentModel!.blocks[0]).toMatchObject({ numbering: { level: 0 } });
	});

	it('lets hosts take over file commands by cancelling the file-command event', () => {
		const editor = mount();
		const seen: string[] = [];
		editor.addEventListener('file-command', (event) => {
			seen.push((event as CustomEvent<{ command: string }>).detail.command);
			event.preventDefault();
		});
		const click = vi.spyOn(HTMLInputElement.prototype, 'click');
		root(editor).querySelector<HTMLButtonElement>('.dve-file-tab')!.click();
		const openPage = [
			...root(editor).querySelectorAll<HTMLButtonElement>('.dve-backstage-nav-item'),
		].find((item) => item.textContent === 'Open')!;
		openPage.click();
		root(editor).querySelector<HTMLButtonElement>('.dve-backstage-primary')!.click();
		expect(seen).toEqual(['open']);
		expect(click).not.toHaveBeenCalled();
		click.mockRestore();
	});

	it('lists compatibility notes in File > Info and from the status bar', () => {
		const editor = mount();
		const notes = root(editor).querySelector<HTMLButtonElement>('.dve-status-notes')!;
		expect(notes.hidden).toBe(false);
		expect(notes.textContent).toBe('1 compatibility note');
		notes.click();
		const backstage = root(editor).querySelector<HTMLElement>('.dve-backstage')!;
		expect(backstage.hidden).toBe(false);
		expect(backstage.textContent).toContain('Pictures are shown as placeholders.');
		expect(backstage.textContent).toContain('Report.docx');
	});

	it('zooms from the status bar and switches to viewing mode from the title bar', () => {
		const editor = mount();
		root(editor).querySelector<HTMLButtonElement>('[aria-label="Zoom in"]')!.click();
		expect(root(editor).querySelector('.dve-zoom-percent')?.textContent).toBe('110%');
		const mode = root(editor).querySelector<HTMLSelectElement>('.dve-mode-select')!;
		mode.value = 'viewing';
		mode.dispatchEvent(new Event('change'));
		expect(editor.readOnly).toBe(true);
	});
});
