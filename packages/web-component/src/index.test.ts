// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { TextSelection } from 'prosemirror-state';
import { loadDocument } from '@christophervr/docx-document';
import { DocxEditorElement, registerDocxEditor } from './index';

vi.mock('@christophervr/docx-document', () => ({ loadDocument: vi.fn() }));
const loadMock = vi.mocked(loadDocument);

describe('DocxEditorElement', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.resetAllMocks();
	});

	it('starts with a usable empty document', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		expect(editor.documentModel?.blocks).toMatchObject([
			{ type: 'paragraph', runs: [{ text: '' }] },
		]);
		expect(editor.shadowRoot?.querySelector('.ProseMirror')).not.toBeNull();
	});

	it('registers idempotently and mounts an accessible, model-backed editor', () => {
		registerDocxEditor();
		registerDocxEditor();
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks[0].type === 'paragraph' &&
			(model.blocks[0].runs[0].text = 'A considered first draft.');
		editor.documentModel = model;
		document.body.append(editor);
		expect(editor.shadowRoot?.querySelector('[role="toolbar"]')?.getAttribute('aria-label')).toBe(
			'Document formatting',
		);
		expect(editor.shadowRoot?.querySelector('.ProseMirror p')?.textContent).toBe(
			'A considered first draft.',
		);
		editor.readOnly = true;
		expect(
			editor.shadowRoot?.querySelector<HTMLButtonElement>('[aria-label="Bold"]')?.disabled,
		).toBe(true);
		expect(editor.shadowRoot?.querySelector<HTMLButtonElement>('[role="tab"]')?.disabled).toBe(
			false,
		);
		expect(
			editor.shadowRoot?.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')?.disabled,
		).toBe(false);
		expect(editor.shadowRoot?.querySelector('.ProseMirror')?.getAttribute('contenteditable')).toBe(
			'false',
		);
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		const before = view.state.doc.toJSON();
		editor.shadowRoot?.querySelector<HTMLButtonElement>('[aria-label="Bold"]')?.click();
		expect(view.state.doc.toJSON()).toEqual(before);
		editor.remove();
	});

	it('reflects editing transactions back into the shared model', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = createDocument();
		document.body.append(editor);
		const changed = vi.fn();
		editor.addEventListener('document-change', changed);
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('New words'));
		expect(editor.documentModel?.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'New words' }],
		});
		expect(changed).toHaveBeenCalledTimes(1);
		expect(changed.mock.calls[0][0].detail).toBe(editor.documentModel);
		editor.remove();
	});

	it('keeps paragraph ids unique when Enter splits a paragraph', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks[0] = { type: 'paragraph', id: 'original', runs: [{ text: 'hello' }] };
		editor.documentModel = model;
		document.body.append(editor);
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.split(3));
		const paragraphs =
			editor.documentModel?.blocks.filter((block) => block.type === 'paragraph') || [];
		expect(paragraphs).toHaveLength(2);
		expect(new Set(paragraphs.map((block) => block.id)).size).toBe(2);
		expect(paragraphs.map((block) => block.id)).toContain('original');
	});

	it('applies selection formatting and preserves run font metadata while editing', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'formatted',
			runs: [{ text: 'select me', fontFamily: 'Garamond', fontSize: 13, color: '#176b62' }],
		};
		editor.documentModel = model;
		document.body.append(editor);
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7)));
		editor.shadowRoot?.querySelector<HTMLButtonElement>('[aria-label="Bold"]')?.click();
		expect(
			editor.shadowRoot?.querySelector('[aria-label="Bold"]')?.getAttribute('aria-pressed'),
		).toBe('true');
		const paragraph = editor.documentModel?.blocks[0];
		expect(paragraph?.type).toBe('paragraph');
		if (paragraph?.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.align).toBeUndefined();
		expect(paragraph.runs[0]).toMatchObject({
			text: 'select',
			bold: true,
			fontFamily: 'Garamond',
			fontSize: 13,
			color: '#176b62',
		});
		expect(paragraph.runs.at(-1)).toMatchObject({
			text: ' me',
			fontFamily: 'Garamond',
			fontSize: 13,
			color: '#176b62',
		});
	});

	it('applies font, color, alignment, table, and page ribbon actions to the shared model', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'ribbon',
				runs: [{ text: 'format me', fontFamily: 'Garamond', fontSize: 13, color: '#176b62' }],
			},
			{ type: 'paragraph', id: 'dve-cell-1', runs: [{ text: 'existing cell id' }] },
			{ type: 'paragraph', id: 'dve-cell-2', runs: [{ text: 'another existing cell id' }] },
			{ type: 'paragraph', id: 'dve-table-3', runs: [{ text: 'existing table id' }] },
		];
		editor.documentModel = model;
		document.body.append(editor);
		const root = editor.shadowRoot!;
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7)));
		const font = root.querySelector<HTMLSelectElement>('[aria-label="Font family"]')!;
		font.value = 'Georgia';
		font.dispatchEvent(new Event('change', { bubbles: true }));
		const color = root.querySelector<HTMLSelectElement>('[aria-label="Font color"]')!;
		color.value = '#c00000';
		color.dispatchEvent(new Event('change', { bubbles: true }));
		root.querySelector<HTMLButtonElement>('[aria-label="Align center"]')!.click();
		const paragraph = editor.documentModel.blocks[0];
		expect(paragraph).toMatchObject({
			type: 'paragraph',
			align: 'center',
			runs: [
				{ text: 'format', fontFamily: 'Georgia', fontSize: 13, color: '#c00000' },
				{ text: ' me', fontFamily: 'Garamond', fontSize: 13, color: '#176b62' },
			],
		});
		root.querySelector<HTMLButtonElement>('[aria-label="Insert table"]')!.click();
		const ids: string[] = [];
		for (const block of editor.documentModel.blocks) {
			if (block.type === 'paragraph') ids.push(block.id);
			else {
				ids.push(block.id);
				for (const row of block.rows)
					for (const cell of row) for (const para of cell.paragraphs) ids.push(para.id);
			}
		}
		expect(new Set(ids).size).toBe(ids.length);
		expect(
			editor.documentModel.blocks.some(
				(block) => block.type === 'table' && block.rows.length === 2 && block.rows[0].length === 2,
			),
		).toBe(true);
		const orientation = root.querySelector<HTMLSelectElement>('[aria-label="Orientation"]')!;
		orientation.value = 'landscape';
		orientation.dispatchEvent(new Event('change', { bubbles: true }));
		expect(editor.documentModel.page.width).toBeGreaterThan(editor.documentModel.page.height);
		root.querySelector<HTMLButtonElement>('[aria-label="Undo"]')!.click();
		expect(editor.documentModel.page.width).toBeLessThan(editor.documentModel.page.height);
		editor.remove();
	});

	it('edits paragraph indentation and spacing through shared model attributes', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'paragraph-format',
			runs: [{ text: 'Format this paragraph' }],
			lineSpacingTwips: 360,
			lineSpacingRule: 'atLeast',
			firstLineTwips: 240,
		};
		editor.documentModel = model;
		document.body.append(editor);
		const root = editor.shadowRoot!;
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 8)));
		root.querySelector<HTMLButtonElement>('[aria-label="Increase indent"]')!.click();
		const spacing = root.querySelector<HTMLSelectElement>('[aria-label="Spacing after"]')!;
		spacing.value = '240';
		spacing.dispatchEvent(new Event('change', { bubbles: true }));
		expect(editor.documentModel.blocks[0]).toMatchObject({
			type: 'paragraph',
			indentLeftTwips: 360,
			spacingAfterTwips: 240,
			lineSpacingTwips: 360,
			lineSpacingRule: 'atLeast',
			firstLineTwips: 240,
		});
		expect(root.querySelector('.ProseMirror p')?.getAttribute('style')).toContain(
			'margin-left: 24px',
		);
		expect(root.querySelector('.ProseMirror p')?.getAttribute('style')).toContain(
			'line-height: max(1.35em, 24px)',
		);
		editor.remove();
	});

	it('exposes linked tab panels and reflows zoom without transform clipping', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		const root = editor.shadowRoot!;
		const tablist = root.querySelector('[role="tablist"]')!;
		const tab = root.querySelector<HTMLButtonElement>(
			'[role="tab"][aria-controls="dve-panel-view"]',
		)!;
		const panel = root.querySelector<HTMLElement>('[id="dve-panel-view"]')!;
		expect(tablist).not.toBeNull();
		expect(panel.getAttribute('aria-labelledby')).toBe(tab.id);
		tab.click();
		expect(tab.getAttribute('aria-selected')).toBe('true');
		expect(panel.hidden).toBe(false);
		const zoom = root.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!;
		zoom.value = '150';
		zoom.dispatchEvent(new Event('change', { bubbles: true }));
		const paper = root.querySelector<HTMLElement>('.dve-paper')!;
		expect(paper.style.getPropertyValue('--dve-zoom')).toBe('1.5');
		expect(paper.style.transform).toBe('none');
		editor.remove();
	});

	it('retains the exact run structure of unchanged paragraphs', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'untouched',
				runs: [
					{ text: 'same ', bold: true },
					{ text: 'style', bold: true },
				],
			},
			{ type: 'paragraph', id: 'edited', runs: [{ text: 'change me' }] },
		];
		editor.documentModel = model;
		document.body.append(editor);
		const before = model.blocks[0];
		const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
		view.dispatch(view.state.tr.insertText('!', view.state.doc.content.size - 1));
		expect(editor.documentModel?.blocks[0]).toBe(before);
		expect(
			(editor.documentModel?.blocks[0] as Extract<typeof before, { type: 'paragraph' }>).runs,
		).toEqual([
			{ text: 'same ', bold: true },
			{ text: 'style', bold: true },
		]);
	});

	it('survives disconnect and reconnect without losing the model', () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		const model = createDocument();
		if (model.blocks[0].type === 'paragraph') model.blocks[0].runs[0].text = 'Keep this text';
		editor.documentModel = model;
		document.body.append(editor);
		editor.remove();
		document.body.append(editor);
		expect(editor.shadowRoot?.querySelector('.ProseMirror p')?.textContent).toBe('Keep this text');
		expect(editor.documentModel?.blocks).toBe(model.blocks);
	});

	it('ignores stale loads and preserves the active imported session for save', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		let resolveOld!: (value: Awaited<ReturnType<typeof loadDocument>>) => void;
		const oldLoad = new Promise<Awaited<ReturnType<typeof loadDocument>>>((resolve) => {
			resolveOld = resolve;
		});
		const current = createDocument();
		if (current.blocks[0].type === 'paragraph') current.blocks[0].runs[0].text = 'current';
		const bytes = new Uint8Array([9, 8, 7]);
		const save = vi.fn().mockResolvedValue(bytes);
		loadMock.mockReturnValueOnce(oldLoad).mockResolvedValueOnce({ model: current, save });
		const first = editor.load(new Uint8Array([1]));
		await editor.load(new Uint8Array([2]));
		const stale = createDocument();
		if (stale.blocks[0].type === 'paragraph') stale.blocks[0].runs[0].text = 'stale';
		resolveOld({ model: stale, save: vi.fn() });
		await first;
		expect(editor.documentModel?.blocks[0]).toMatchObject({ runs: [{ text: 'current' }] });
		await expect(editor.save()).resolves.toBe(bytes);
		expect(save).toHaveBeenCalledWith(current);
	});

	it('exposes import errors as events while rejecting the load promise', async () => {
		const editor = new DocxEditorElement();
		const error = new Error('bad package');
		loadMock.mockRejectedValueOnce(error);
		const handler = vi.fn();
		editor.addEventListener('document-error', handler);
		await expect(editor.load(new Uint8Array())).rejects.toBeTruthy();
		expect(handler).toHaveBeenCalledTimes(1);
		expect(handler.mock.calls[0][0].detail).toBeInstanceOf(Error);
		expect(handler.mock.calls[0][0].detail).toBe(error);
	});
});
