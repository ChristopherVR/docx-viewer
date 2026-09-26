import './style.css';
import { createDocument, saveDocx, type DocumentModel } from '@christophervr/docx-core';
import { detectDocumentFormat } from '@christophervr/docx-document';
import { mountFramework } from './framework';
const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const themeKey = 'vitepress-theme-appearance';
const savedTheme = localStorage.getItem(themeKey);
const applyTheme = (theme: 'light' | 'dark') => {
	document.documentElement.dataset.theme = theme;
	const toggle = get<HTMLButtonElement>('theme-toggle');
	const dark = theme === 'dark';
	toggle.textContent = dark ? 'Light' : 'Dark';
	toggle.setAttribute('aria-label', `Switch to ${dark ? 'light' : 'dark'} theme`);
	toggle.setAttribute('aria-pressed', String(dark));
};
applyTheme(
	savedTheme === 'dark' || savedTheme === 'light'
		? savedTheme
		: matchMedia('(prefers-color-scheme: dark)').matches
			? 'dark'
			: 'light',
);
get<HTMLButtonElement>('theme-toggle').addEventListener('click', () => {
	const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
	localStorage.setItem(themeKey, next);
	applyTheme(next);
});
const sample = createDocument();
sample.blocks = [
	{
		type: 'paragraph',
		id: 'title',
		runs: [{ text: 'Document title', fontFamily: 'Calibri', fontSize: 26, bold: true }],
	},
	{
		type: 'paragraph',
		id: 'subtitle',
		runs: [{ text: 'Add a subtitle or date here', fontSize: 10, color: '#666666' }],
	},
	{
		type: 'paragraph',
		id: 'intro',
		runs: [
			{
				text: 'Start with your first paragraph. Use the ribbon to set type, size, alignment, and page options as you work.',
			},
		],
	},
	{
		type: 'paragraph',
		id: 'body',
		runs: [
			{ text: 'A document workspace. ', bold: true },
			{
				text: 'Select text to format it, insert a simple table, open a DOCX file, or start a new document. The shared editing surface works across supported framework bindings.',
			},
		],
	},
	{
		type: 'paragraph',
		id: 'quote',
		align: 'center',
		runs: [{ text: 'Clarity is a kindness to the reader.', italic: true, fontSize: 16 }],
	},
	{
		type: 'paragraph',
		id: 'heading',
		runs: [{ text: 'A small plan for a good first draft', bold: true }],
	},
	{
		type: 'table',
		id: 'table',
		rows: [
			[
				{
					paragraphs: [{ type: 'paragraph', id: 'h1', runs: [{ text: 'START WITH', bold: true }] }],
				},
				{
					paragraphs: [
						{ type: 'paragraph', id: 'h2', runs: [{ text: 'MAKE IT COUNT', bold: true }] },
					],
				},
			],
			[
				{ paragraphs: [{ type: 'paragraph', id: 'c1', runs: [{ text: 'One clear idea' }] }] },
				{
					paragraphs: [
						{ type: 'paragraph', id: 'c2', runs: [{ text: 'Give each paragraph a purpose.' }] },
					],
				},
			],
			[
				{ paragraphs: [{ type: 'paragraph', id: 'c3', runs: [{ text: 'A thoughtful edit' }] }] },
				{
					paragraphs: [
						{ type: 'paragraph', id: 'c4', runs: [{ text: 'Read it once more, then share.' }] },
					],
				},
			],
		],
	},
];
let filename = 'Untitled.docx';
const status = (text: string) => {
	get('status').textContent = text;
};
const showWarnings = (model: DocumentModel) => {
	get('warnings').replaceChildren(
		...model.warnings.map((w) => {
			const li = document.createElement('li');
			li.textContent = w;
			return li;
		}),
	);
};
const editor = await mountFramework(get('editor'), {
	documentModel: sample,
	onDocumentChange(model) {
		get('state').textContent = 'Unsaved changes';
		showWarnings(model);
	},
	onDocumentError(error) {
		status(error.message);
	},
});
get<HTMLInputElement>('readonly').addEventListener('change', (event) => {
	editor.element.readOnly = (event.target as HTMLInputElement).checked;
});
get<HTMLInputElement>('file').addEventListener('change', async (event) => {
	const input = event.target as HTMLInputElement;
	const file = input.files?.[0];
	if (!file) return;
	try {
		status('Opening document…');
		const bytes = new Uint8Array(await file.arrayBuffer());
		const format = detectDocumentFormat(bytes);
		await editor.load(bytes);
		filename = `${file.name.replace(/\.[^.]+$/, '')}.${format}`;
		get('filename').textContent = filename;
		get('state').textContent = 'Ready to edit';
		if (editor.element.documentModel) showWarnings(editor.element.documentModel);
		status(`Opened ${filename}. Review format support before editing complex documents.`);
	} catch (error) {
		status(error instanceof Error ? error.message : String(error));
	} finally {
		input.value = '';
	}
});
get('new').addEventListener('click', () => {
	if (
		get('state').textContent === 'Unsaved changes' &&
		!confirm('Discard unsaved changes and start a new document?')
	)
		return;
	editor.element.documentModel = createDocument();
	filename = 'Untitled.docx';
	get('filename').textContent = filename;
	get('state').textContent = 'Ready to edit';
	get('warnings').replaceChildren();
	status('New document ready.');
});
function download(bytes: Uint8Array, name: string) {
	const url = URL.createObjectURL(
		new Blob([new Uint8Array(bytes)], {
			type: name.endsWith('.doc')
				? 'application/msword'
				: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		}),
	);
	const anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = name;
	anchor.click();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
}
get('save').addEventListener('click', async () => {
	try {
		download(await editor.save(), filename);
		get('state').textContent = 'Saved';
		status(`Saved ${filename}.`);
	} catch (error) {
		status(error instanceof Error ? error.message : String(error));
	}
});
get('export-docx').addEventListener('click', async () => {
	try {
		if (!editor.element.documentModel) return;
		download(
			await saveDocx(structuredClone(editor.element.documentModel)),
			filename.replace(/\.[^.]+$/, '') + '.docx',
		);
		status('Exported visible content as a new DOCX document.');
	} catch (error) {
		status(error instanceof Error ? error.message : String(error));
	}
});
