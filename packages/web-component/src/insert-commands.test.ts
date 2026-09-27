// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	loadDocx,
	type DocumentModel,
	type Paragraph,
} from '@christophervr/docx-core';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement, registerDocxEditor } from './index';
import { stagePicture, insertPicture } from './picture-commands';
import { followLinkAt } from './link-commands';

registerDocxEditor();

const PNG = Uint8Array.from(
	atob(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
	),
	(char) => char.charCodeAt(0),
);

function mount(model: DocumentModel): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.documentModel = model;
	document.body.append(editor);
	return editor;
}
const view = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const shadow = (editor: DocxEditorElement) => editor.shadowRoot!;
const firstParagraph = (editor: DocxEditorElement) => editor.documentModel!.blocks[0] as Paragraph;
function textModel(text: string): DocumentModel {
	const model = createDocument();
	model.blocks[0] = { type: 'paragraph', id: 'p1', runs: [{ text }] };
	return model;
}
function select(editor: DocxEditorElement, from: number, to: number) {
	const pm = view(editor);
	pm.dispatch(pm.state.tr.setSelection(TextSelection.create(pm.state.doc, from, to)));
}
function ribbon(editor: DocxEditorElement, label: string) {
	shadow(editor).querySelector<HTMLButtonElement>(`.dve-ribbon [aria-label="${label}"]`)!.click();
}
function dialogButton(dialog: HTMLElement, text: string) {
	return [...dialog.querySelectorAll('button')].find((button) => button.textContent === text)!;
}

describe('editor insert and formatting commands', () => {
	afterEach(() => document.body.replaceChildren());

	it('inserts a picture and writes its media part, relationship and drawing on save', async () => {
		const editor = mount(textModel('Before'));
		const staged = await stagePicture(new File([PNG], 'Logo.png', { type: 'image/png' }), 600);
		const inserts = (editor as unknown as { inserts: { pendingMedia: Map<string, unknown> } })
			.inserts;
		inserts.pendingMedia.set(staged.image.partName, staged.media);
		select(editor, 7, 7);
		insertPicture(view(editor), staged.image);
		expect(firstParagraph(editor).runs.at(-1)?.image?.partName).toBe(staged.image.partName);
		const saved = await editor.save();
		const zip = await JSZip.loadAsync(saved);
		expect(await zip.file(staged.image.partName)!.async('uint8array')).toEqual(PNG);
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(rels).toContain(staged.image.partName.replace('word/', ''));
		const reloaded = await loadDocx(saved);
		const image = (reloaded.model.blocks[0] as Paragraph).runs.find((run) => run.image)?.image;
		expect(image).toMatchObject({ contentType: 'image/png', altText: 'Logo' });
		expect(reloaded.media?.get(image!.partName)).toEqual(PNG);
	});

	it('rejects unsupported picture formats', async () => {
		await expect(
			stagePicture(new File(['<svg/>'], 'art.svg', { type: 'image/svg+xml' }), 600),
		).rejects.toThrow(/PNG, JPEG, GIF or BMP/);
	});

	it('adds, edits and removes a hyperlink through the link dialog', () => {
		const editor = mount(textModel('Visit our site today'));
		select(editor, 11, 15);
		ribbon(editor, 'Insert link');
		const dialog = shadow(editor).querySelector<HTMLElement>('.dve-link-dialog')!;
		expect(dialog.hidden).toBe(false);
		dialog.querySelector<HTMLInputElement>('[aria-label="Address"]')!.value =
			'https://example.com/';
		dialog.querySelector<HTMLInputElement>('[aria-label="ScreenTip"]')!.value = 'Example';
		dialogButton(dialog, 'Insert').click();
		const linked = firstParagraph(editor).runs.find((run) => run.link);
		expect(linked).toMatchObject({
			text: 'site',
			link: { href: 'https://example.com/', tooltip: 'Example' },
		});
		select(editor, 12, 12);
		ribbon(editor, 'Insert link');
		expect(dialog.querySelector<HTMLInputElement>('[aria-label="Address"]')!.value).toBe(
			'https://example.com/',
		);
		dialogButton(dialog, 'Remove link').click();
		expect(firstParagraph(editor).runs.some((run) => run.link)).toBe(false);
	});

	it('refuses script links and reports the error', () => {
		const editor = mount(textModel('Click here'));
		const errors: string[] = [];
		editor.addEventListener('document-error', (event) =>
			errors.push((event as CustomEvent<Error>).detail.message),
		);
		select(editor, 1, 6);
		ribbon(editor, 'Insert link');
		const dialog = shadow(editor).querySelector<HTMLElement>('.dve-link-dialog')!;
		dialog.querySelector<HTMLInputElement>('[aria-label="Address"]')!.value = 'javascript:alert(1)';
		dialogButton(dialog, 'Insert').click();
		expect(errors[0]).toMatch(/http:\/\/, https:\/\/ or mailto:/);
		expect(firstParagraph(editor).runs.some((run) => run.link)).toBe(false);
	});

	it('follows an internal link to its bookmark', () => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'p1', runs: [{ text: 'Jump', link: { anchor: 'Target' } }] },
			{ type: 'paragraph', id: 'p2', runs: [{ text: 'Destination' }], bookmarks: ['Target'] },
		];
		const editor = mount(model);
		let scrolled = '';
		HTMLElement.prototype.scrollIntoView = function () {
			scrolled = this.textContent ?? '';
		};
		expect(followLinkAt(view(editor), 2)).toBe(true);
		expect(scrolled).toBe('Destination');
	});

	it('applies a character style and renders its inherited formatting', () => {
		const model = textModel('Styled words');
		model.characterStyles = {
			docDefaults: {},
			styles: {
				Strong: { id: 'Strong', type: 'character', name: 'Strong', formatting: { bold: true } },
			},
			warnings: [],
		};
		const editor = mount(model);
		const picker = shadow(editor).querySelector<HTMLSelectElement>('[data-character-styles]')!;
		expect([...picker.options].map((option) => option.textContent)).toEqual([
			'No character style',
			'Strong',
		]);
		select(editor, 1, 7);
		picker.value = 'Strong';
		picker.dispatchEvent(new Event('change'));
		expect(firstParagraph(editor).runs[0]).toMatchObject({ text: 'Styled', style: 'Strong' });
		const styled = shadow(editor).querySelector<HTMLElement>('span[data-docx-style="Strong"]')!;
		// Inline decorations render inside mark elements.
		const decorated = styled.querySelector<HTMLElement>('[style*="font-weight"]');
		expect(decorated?.style.fontWeight).toBe('700');
	});

	it('hides hidden text until "Show hidden text" is on', () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [{ text: 'Visible ' }, { text: 'secret', vanish: true }],
		};
		const editor = mount(model);
		const paper = shadow(editor).querySelector<HTMLElement>('.dve-paper')!;
		expect(paper.querySelector('.dve-hidden-text')?.textContent).toBe('secret');
		expect(paper.hasAttribute('data-show-hidden')).toBe(false);
		ribbon(editor, 'Show hidden text');
		expect(paper.hasAttribute('data-show-hidden')).toBe(true);
	});
});

describe('toggle property display', () => {
	afterEach(() => document.body.replaceChildren());

	it('cancels direct bold on a run whose character style is also bold, as Word does', () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [{ text: 'Plain again', bold: true, style: 'Strong' }],
		};
		model.characterStyles = {
			docDefaults: {},
			styles: { Strong: { id: 'Strong', type: 'character', formatting: { bold: true } } },
			warnings: [],
		};
		const editor = mount(model);
		const decorated = shadow(editor).querySelector<HTMLElement>(
			'.dve-paper strong [style*="font-weight"]',
		);
		expect(decorated?.style.fontWeight).toBe('400');
	});
});
