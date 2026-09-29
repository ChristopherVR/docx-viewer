import type { DocumentModel } from '@christophervr/docx-core';
import { icon } from './chrome-icons';
import type { DocumentStats } from './document-stats';
import type { FileCommand } from './file-commands';
import { pageSizeOf } from './page-size';
import { sectionsOf } from './section-commands';

/** What the File view needs from the editor; every value is read when a page opens. */
export interface BackstageHandlers {
	fileCommand(command: FileCommand, fileName?: string): void;
	close(): void;
	summary(): {
		fileName: string;
		model: DocumentModel;
		words: number;
		stats: DocumentStats;
		saveState: string;
	};
	options(): { locale: string; theme: string; author: string };
	setOption(key: 'locale' | 'theme' | 'author', value: string): void;
}

export interface PageContext {
	handlers: BackstageHandlers;
	t(text: string): string;
	content: HTMLElement;
}

const heading = (text: string, level: 'h2' | 'h3' = 'h2') => {
	const element = document.createElement(level);
	element.textContent = text;
	return element;
};
const paragraph = (text: string, className = '') => {
	const element = document.createElement('p');
	element.textContent = text;
	if (className) element.className = className;
	return element;
};
const primary = (label: string, onClick: () => void) => {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'dve-backstage-primary';
	button.textContent = label;
	button.addEventListener('click', onClick);
	return button;
};
const labelled = (label: string, control: HTMLElement) => {
	const wrap = document.createElement('label');
	wrap.className = 'dve-backstage-field';
	const text = document.createElement('span');
	text.textContent = label;
	control.setAttribute('aria-label', label);
	wrap.append(text, control);
	return wrap;
};
/** A Word-style action card: a title, a description and one button. */
function card(title: string, description: string, action: HTMLElement): HTMLElement {
	const el = document.createElement('div');
	el.className = 'dve-backstage-card';
	const text = document.createElement('div');
	text.append(heading(title, 'h3'), paragraph(description, 'dve-backstage-muted'));
	el.append(text, action);
	return el;
}
const facts = (rows: Array<[string, string]>, t: (text: string) => string) => {
	const list = document.createElement('dl');
	list.className = 'dve-backstage-facts';
	for (const [label, value] of rows) {
		const term = document.createElement('dt');
		term.textContent = t(label);
		const detail = document.createElement('dd');
		detail.textContent = value;
		list.append(term, detail);
	}
	return list;
};

export function renderInfo({ handlers, t, content }: PageContext): void {
	const { fileName, model, stats, words, saveState } = handlers.summary();
	const section = sectionsOf(model)[0];
	const size = section ? pageSizeOf(section) : null;
	const status =
		saveState === 'dirty'
			? 'Unsaved changes'
			: saveState === 'saved-local'
				? 'Saved to this device'
				: 'Saved';
	const title = heading(fileName);
	title.classList.add('dve-backstage-title');
	const notes = document.createElement('ul');
	notes.className = 'dve-compatibility-notes';
	for (const warning of model.warnings) {
		const item = document.createElement('li');
		const text = document.createElement('span');
		text.textContent = warning;
		item.append(icon('warning'), text);
		notes.append(item);
	}
	if (!model.warnings.length) {
		const item = document.createElement('li');
		item.className = 'dve-compatibility-empty';
		item.textContent = t('No compatibility notes for this document.');
		notes.append(item);
	}
	const left = document.createElement('div');
	left.append(
		title,
		paragraph(t(status), 'dve-backstage-muted'),
		heading(t('Inspect document'), 'h3'),
		paragraph(t('Compatibility notes'), 'dve-backstage-muted'),
		notes,
		heading(t('Manage document'), 'h3'),
		facts(
			[
				['Track changes', model.trackChanges ? t('On') : t('Off')],
				['Tracked changes', String(stats.revisions)],
				['Comments', String(stats.comments)],
			],
			t,
		),
	);
	const right = document.createElement('aside');
	right.className = 'dve-backstage-properties';
	right.append(
		heading(t('Properties'), 'h3'),
		facts(
			[
				['Words', String(words)],
				['Characters', String(stats.characters)],
				['Characters (with spaces)', String(stats.charactersWithSpaces)],
				['Paragraphs', String(stats.paragraphs)],
				['Tables', String(stats.tables)],
				['Sections', String(stats.sections)],
				['Page size', size ? size.toUpperCase() : t('Custom')],
				['Orientation', section?.orientation === 'landscape' ? t('Landscape') : t('Portrait')],
			],
			t,
		),
	);
	const layout = document.createElement('div');
	layout.className = 'dve-backstage-columns';
	layout.append(left, right);
	content.replaceChildren(heading(t('Info')), layout);
}

export function renderNew({ handlers, t, content }: PageContext): void {
	const blank = document.createElement('button');
	blank.type = 'button';
	blank.className = 'dve-backstage-tile';
	const preview = document.createElement('span');
	preview.className = 'dve-backstage-tile-preview';
	const label = document.createElement('span');
	label.textContent = t('Blank document');
	blank.append(preview, label);
	blank.setAttribute('aria-label', t('Blank document'));
	blank.addEventListener('click', () => {
		handlers.close();
		handlers.fileCommand('new');
	});
	content.replaceChildren(heading(t('New')), blank);
}

export function renderOpen({ handlers, t, content }: PageContext): void {
	content.replaceChildren(
		heading(t('Open')),
		paragraph(t('Open a Word document (.docx or .doc) from this device.')),
		primary(t('Browse…'), () => {
			handlers.close();
			handlers.fileCommand('open');
		}),
		paragraph(t('Files are processed entirely in the browser.'), 'dve-backstage-muted'),
	);
}

export function renderSaveAs({ handlers, t, content }: PageContext): void {
	const { fileName } = handlers.summary();
	const name = document.createElement('input');
	name.type = 'text';
	name.value = fileName;
	const format = document.createElement('input');
	format.type = 'text';
	format.readOnly = true;
	format.value = fileName.toLowerCase().endsWith('.doc')
		? t('Word 97-2003 Document (*.doc)')
		: t('Word Document (*.docx)');
	const save = () => {
		if (!name.value.trim()) return;
		handlers.close();
		handlers.fileCommand('saveAs', name.value);
	};
	name.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') save();
	});
	content.replaceChildren(
		heading(t('Save As')),
		paragraph(t('Save a copy of this document to your device.'), 'dve-backstage-muted'),
		labelled(t('File name'), name),
		labelled(t('Format'), format),
		primary(t('Save'), save),
	);
}

export function renderPrint({ handlers, t, content }: PageContext): void {
	const { model } = handlers.summary();
	const section = sectionsOf(model)[0];
	const inches = (twips: number) => `${(twips / 1440).toFixed(2).replace(/\.?0+$/, '')}"`;
	content.replaceChildren(
		heading(t('Print')),
		primary(t('Print'), () => {
			handlers.close();
			handlers.fileCommand('print');
		}),
		facts(
			[
				[
					'Paper size',
					section
						? `${pageSizeOf(section)?.toUpperCase() ?? t('Custom')} (${inches(section.pageWidthTwips)} × ${inches(section.pageHeightTwips)})`
						: '',
				],
				['Orientation', section?.orientation === 'landscape' ? t('Landscape') : t('Portrait')],
			],
			t,
		),
		paragraph(
			t(
				'Printing uses the browser print dialog. Page breaks come from Print Layout, which approximates Word.',
			),
			'dve-backstage-muted',
		),
	);
}

export function renderExport({ handlers, t, content }: PageContext): void {
	const run = (command: FileCommand) => () => {
		handlers.close();
		handlers.fileCommand(command);
	};
	content.replaceChildren(
		heading(t('Export')),
		card(
			t('Create a PDF'),
			t('Opens the print dialog; choose Save as PDF as the destination.'),
			primary(t('Create PDF'), run('print')),
		),
		card(
			t('Save a copy as DOCX'),
			t('Download the document as a Word .docx file.'),
			primary(t('Save a copy as DOCX'), run('export')),
		),
		card(
			t('Export as plain text'),
			t('Download the text only, without formatting, pictures or notes.'),
			primary(t('Export as plain text'), run('exportText')),
		),
	);
}

const LANGUAGES: Array<[string, string]> = [
	['en', 'English'],
	['fr', 'Français'],
	['de', 'Deutsch'],
	['es', 'Español'],
	['zh-CN', '简体中文'],
];

export function renderOptions({ handlers, t, content }: PageContext): void {
	const current = handlers.options();
	const language = document.createElement('select');
	for (const [value, label] of LANGUAGES) language.append(new Option(label, value));
	language.value = current.locale;
	language.addEventListener('change', () => handlers.setOption('locale', language.value));
	const theme = document.createElement('select');
	for (const [value, label] of [
		['auto', 'Automatic'],
		['light', 'Light'],
		['dark', 'Dark'],
	] as const)
		theme.append(new Option(t(label), value));
	theme.value = ['light', 'dark'].includes(current.theme) ? current.theme : 'auto';
	theme.addEventListener('change', () => handlers.setOption('theme', theme.value));
	const author = document.createElement('input');
	author.type = 'text';
	author.value = current.author;
	author.addEventListener('change', () => handlers.setOption('author', author.value));
	content.replaceChildren(
		heading(t('Options')),
		heading(t('General'), 'h3'),
		labelled(t('Display language'), language),
		labelled(t('Theme'), theme),
		labelled(t('Author name'), author),
		paragraph(
			t('Options apply to this editor and are not stored between sessions.'),
			'dve-backstage-muted',
		),
	);
}
