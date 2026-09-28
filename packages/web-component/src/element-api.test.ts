// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, loadDocx, type DocumentModel } from '@christophervr/docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';
import { RIBBON_ACTION_IDS } from './ribbon-action-ids';
import { createRibbon } from './ribbon';
import { ribbonControlId } from './ribbon-visibility';

registerDocxEditor();

function paragraphs(count: number): DocumentModel {
	const model = createDocument();
	model.page = {
		width: 300,
		height: 200,
		marginTop: 20,
		marginRight: 20,
		marginBottom: 20,
		marginLeft: 20,
	};
	model.blocks = Array.from({ length: count }, (_, index) => ({
		type: 'paragraph' as const,
		id: `p${index}`,
		runs: [{ text: `Paragraph number ${index} with enough words to fill a line or two.` }],
	}));
	return model;
}

function mount(model?: DocumentModel) {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	if (model) editor.documentModel = model;
	document.body.append(editor);
	return editor;
}

function blobBytes(blob: Blob): Promise<Uint8Array> {
	return new Promise((resolve, reject) => {
		const reader = new FileReader();
		reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
		reader.onerror = () => reject(reader.error);
		reader.readAsArrayBuffer(blob);
	});
}

const shell = (editor: DocxEditorElement) => editor.shadowRoot!;

describe('dirty tracking and save API', () => {
	afterEach(() => document.body.replaceChildren());

	it('is clean at first, dirty after an edit, and clean after markClean()', () => {
		const editor = mount();
		const changes: boolean[] = [];
		editor.addEventListener('dirty-change', (event) => changes.push(event.detail));
		expect(editor.dirty).toBe(false);
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('hello', 1));
		expect(editor.dirty).toBe(true);
		view.dispatch(view.state.tr.insertText('!', 1));
		editor.markClean();
		expect(editor.dirty).toBe(false);
		expect(changes).toEqual([true, false]);
		expect(shell(editor).querySelector('.dve-save-state')?.textContent ?? '').not.toContain(
			'Unsaved',
		);
	});

	it('becomes clean when a document is loaded or replaced', () => {
		const editor = mount();
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('draft', 1));
		expect(editor.dirty).toBe(true);
		editor.documentModel = createDocument();
		expect(editor.dirty).toBe(false);
	});

	it('save() returns a docx Blob that re-parses to the edited model', async () => {
		const editor = mount(paragraphs(2));
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('Edited: ', 1));
		const blob = await editor.save();
		expect(blob).toBeInstanceOf(Blob);
		expect(blob.type).toBe(
			'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		);
		const parsed = await loadDocx(await blobBytes(blob));
		const first = parsed.model.blocks[0];
		expect(first.type === 'paragraph' && first.runs.map((run) => run.text).join('')).toContain(
			'Edited: Paragraph number 0',
		);
		// Saving does not claim the host persisted it.
		expect(editor.dirty).toBe(true);
	});

	it('download() saves through the same path, names the file and marks the document clean', async () => {
		const editor = mount(paragraphs(1));
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('x', 1));
		const urls: Blob[] = [];
		URL.createObjectURL = vi.fn((blob: Blob | MediaSource) => {
			urls.push(blob as Blob);
			return 'blob:test';
		});
		URL.revokeObjectURL = vi.fn();
		const clicked: string[] = [];
		vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(
			function (this: HTMLAnchorElement) {
				clicked.push(this.download);
			},
		);
		await editor.download('report');
		expect(clicked).toEqual(['report.docx']);
		expect(urls).toHaveLength(1);
		expect(editor.dirty).toBe(false);
		vi.restoreAllMocks();
	});

	it('File > Save marks the document clean through the same state', async () => {
		const editor = mount(paragraphs(1));
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('x', 1));
		URL.createObjectURL = vi.fn(() => 'blob:test');
		URL.revokeObjectURL = vi.fn();
		vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => {});
		shell(editor).querySelector<HTMLButtonElement>('[aria-label="Save"]')!.click();
		await vi.waitFor(() => expect(editor.dirty).toBe(false));
		vi.restoreAllMocks();
	});
});

describe('UI customisation', () => {
	afterEach(() => document.body.replaceChildren());

	it('hiddenActions hides the named ribbon controls, empty groups and tabs', () => {
		const editor = mount();
		editor.hiddenActions = [
			'Bold',
			'Print',
			'Show hidden text',
			'Page thumbnails',
			'Zoom',
			'Layout view',
		];
		const hidden = (label: string) =>
			shell(editor).querySelector(`[aria-label="${label}"]`)!.hasAttribute('data-dve-hidden');
		expect(hidden('Bold')).toBe(true);
		expect(hidden('Italic')).toBe(false);
		expect(shell(editor).querySelector('#dve-tab-view')!.hasAttribute('data-dve-hidden')).toBe(
			true,
		);
		expect(shell(editor).querySelector('#dve-tab-home')!.hasAttribute('data-dve-hidden')).toBe(
			false,
		);
		editor.hiddenActions = [];
		expect(hidden('Bold')).toBe(false);
		expect(shell(editor).querySelector('#dve-tab-view')!.hasAttribute('data-dve-hidden')).toBe(
			false,
		);
	});

	it('applies hiddenActions set before the element connects', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.hiddenActions = ['Italic'];
		document.body.append(editor);
		expect(
			shell(editor).querySelector('[aria-label="Italic"]')!.hasAttribute('data-dve-hidden'),
		).toBe(true);
	});

	it('showToolbar and show-toolbar hide the ribbon and stay in sync', () => {
		const editor = mount();
		const ribbon = () => shell(editor).querySelector('.dve-ribbon')!;
		expect(editor.showToolbar).toBe(true);
		editor.showToolbar = false;
		expect(ribbon().hasAttribute('data-dve-hidden')).toBe(true);
		expect(editor.getAttribute('show-toolbar')).toBe('false');
		editor.setAttribute('show-toolbar', 'true');
		expect(editor.showToolbar).toBe(true);
		expect(ribbon().hasAttribute('data-dve-hidden')).toBe(false);
		editor.setAttribute('show-toolbar', 'false');
		expect(editor.showToolbar).toBe(false);
		editor.removeAttribute('show-toolbar');
		expect(editor.showToolbar).toBe(true);
	});

	it('showThumbnails opens the rail, reflects to show-thumbnails and follows the View toggle', () => {
		const editor = mount(paragraphs(2));
		const rail = () => shell(editor).querySelector<HTMLElement>('.dve-pages-rail')!;
		expect(editor.showThumbnails).toBe(false);
		expect(rail().hidden).toBe(true);
		editor.showThumbnails = true;
		expect(rail().hidden).toBe(false);
		expect(editor.hasAttribute('show-thumbnails')).toBe(true);
		// Web/draft view: an honest explanation, no page list.
		expect(rail().querySelector('.dve-pages-message')!.textContent).toContain('Print Layout');
		const toggle = shell(editor).querySelector<HTMLButtonElement>(
			'[aria-label="Page thumbnails"]',
		)!;
		expect(toggle.getAttribute('aria-pressed')).toBe('true');
		toggle.click();
		expect(editor.showThumbnails).toBe(false);
		expect(rail().hidden).toBe(true);
		editor.setAttribute('show-thumbnails', '');
		expect(editor.showThumbnails).toBe(true);
	});
});

describe('page rail in Print Layout', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.useRealTimers();
	});

	it('lists a thumbnail per laid-out page and emits page-change', () => {
		vi.useFakeTimers();
		const editor = mount(paragraphs(40));
		const events: { page: number; pageCount: number }[] = [];
		editor.addEventListener('page-change', (event) => events.push(event.detail));
		editor.showThumbnails = true;
		const view = shell(editor).querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!;
		view.value = 'print';
		view.dispatchEvent(new Event('change'));
		vi.advanceTimersByTime(500);
		const pages = shell(editor).querySelectorAll('.dve-canvas .dve-print-page').length;
		expect(pages).toBeGreaterThan(1);
		const options = shell(editor).querySelectorAll('.dve-pages-rail [role="option"]');
		expect(options).toHaveLength(pages);
		expect(events.at(-1)).toEqual({ page: 1, pageCount: pages });
		expect(shell(editor).querySelector('.dve-status-page')!.textContent).toBe(`Page 1 of ${pages}`);
		expect(shell(editor).querySelector<HTMLElement>('.dve-status-page')!.title).toContain(
			'approximation',
		);
	});
});

describe('ribbon action ids', () => {
	it('lists exactly the labelled ribbon controls', () => {
		const ribbon = createRibbon('en');
		const ids = new Set<string>();
		for (const control of ribbon.querySelectorAll<HTMLElement>(
			'.ribbon-panel button[aria-label], .ribbon-panel select[aria-label], .ribbon-panel input[aria-label]',
		))
			ids.add(ribbonControlId(control)!);
		expect([...ids].sort()).toEqual([...RIBBON_ACTION_IDS].sort());
	});
});
