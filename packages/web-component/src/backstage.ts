import type { DocumentModel } from '@christophervr/docx-core';
import { icon, type ChromeIcon } from './chrome-icons';
import type { FileCommand } from './file-commands';
import { translateUiText } from './localization';

export type BackstagePage = 'info' | 'new' | 'open';

export interface BackstageHandlers {
	fileCommand(command: FileCommand): void;
	close(): void;
	/** Current document facts for the Info page. */
	summary(): { fileName: string; model: DocumentModel; words: number };
}

export interface Backstage {
	element: HTMLElement;
	open(page?: BackstagePage): void;
	close(): void;
	readonly isOpen: boolean;
}

function navButton(iconName: ChromeIcon, label: string): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'dve-backstage-nav-item';
	const text = document.createElement('span');
	text.textContent = label;
	button.append(icon(iconName), text);
	return button;
}

function heading(text: string): HTMLHeadingElement {
	const element = document.createElement('h2');
	element.textContent = text;
	return element;
}

function countBlocks(model: DocumentModel): { paragraphs: number; tables: number } {
	let paragraphs = 0;
	let tables = 0;
	for (const block of model.blocks) {
		if (block.type === 'paragraph') paragraphs++;
		else {
			tables++;
			for (const row of block.rows) for (const cell of row) paragraphs += cell.paragraphs.length;
		}
	}
	return { paragraphs, tables };
}

/** Word's File view: Info (properties and compatibility notes), New, Open, Save, copy, Print. */
export function createBackstage(handlers: BackstageHandlers): Backstage {
	const element = document.createElement('section');
	element.className = 'dve-backstage';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-modal', 'true');
	element.hidden = true;
	const nav = document.createElement('nav');
	nav.className = 'dve-backstage-nav';
	const content = document.createElement('div');
	content.className = 'dve-backstage-content';
	const t = (text: string) => translateUiText(element, text);
	element.setAttribute('aria-label', t('File'));

	const back = navButton('back', 'Back to document');
	back.classList.add('dve-backstage-back');
	back.addEventListener('click', () => handlers.close());
	const pages: Record<BackstagePage, HTMLButtonElement> = {
		info: navButton('info', 'Info'),
		new: navButton('file', 'New'),
		open: navButton('folder', 'Open'),
	};
	const commands: [ChromeIcon, string, FileCommand][] = [
		['save', 'Save', 'save'],
		['copy', 'Save a copy as DOCX', 'export'],
		['print', 'Print', 'print'],
	];
	nav.append(back, pages.info, pages.new, pages.open);
	for (const [iconName, label, command] of commands) {
		const button = navButton(iconName, label);
		button.addEventListener('click', () => {
			handlers.close();
			handlers.fileCommand(command);
		});
		nav.append(button);
	}

	const renderInfo = () => {
		const { fileName, model, words } = handlers.summary();
		const { paragraphs, tables } = countBlocks(model);
		const facts = document.createElement('dl');
		facts.className = 'dve-backstage-facts';
		const rows: [string, string][] = [
			['File name', fileName],
			['Words', String(words)],
			['Paragraphs', String(paragraphs)],
			['Tables', String(tables)],
		];
		for (const [label, value] of rows) {
			const term = document.createElement('dt');
			term.textContent = t(label);
			const detail = document.createElement('dd');
			detail.textContent = value;
			facts.append(term, detail);
		}
		const notesHeading = document.createElement('h3');
		notesHeading.textContent = t('Compatibility notes');
		const notes = document.createElement('ul');
		notes.className = 'dve-compatibility-notes';
		for (const warning of model.warnings) {
			const item = document.createElement('li');
			item.append(icon('warning'));
			const text = document.createElement('span');
			text.textContent = warning;
			item.append(text);
			notes.append(item);
		}
		if (!model.warnings.length) {
			const item = document.createElement('li');
			item.className = 'dve-compatibility-empty';
			item.textContent = t('No compatibility notes for this document.');
			notes.append(item);
		}
		const propertiesHeading = document.createElement('h3');
		propertiesHeading.textContent = t('Document properties');
		content.replaceChildren(heading(t('Info')), propertiesHeading, facts, notesHeading, notes);
	};
	const renderNew = () => {
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
	};
	const renderOpen = () => {
		const description = document.createElement('p');
		description.textContent = t('Open a Word document (.docx or .doc) from this device.');
		const browse = document.createElement('button');
		browse.type = 'button';
		browse.className = 'dve-backstage-primary';
		browse.textContent = t('Browse…');
		browse.addEventListener('click', () => {
			handlers.close();
			handlers.fileCommand('open');
		});
		const privacy = document.createElement('p');
		privacy.className = 'dve-backstage-muted';
		privacy.textContent = t('Files are processed entirely in the browser.');
		content.replaceChildren(heading(t('Open')), description, browse, privacy);
	};
	const renderers: Record<BackstagePage, () => void> = {
		info: renderInfo,
		new: renderNew,
		open: renderOpen,
	};
	const show = (page: BackstagePage) => {
		for (const [key, button] of Object.entries(pages))
			button.setAttribute('aria-current', String(key === page));
		renderers[page]();
	};
	for (const [key, button] of Object.entries(pages) as [BackstagePage, HTMLButtonElement][])
		button.addEventListener('click', () => show(key));
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') handlers.close();
	});
	element.append(nav, content);
	return {
		element,
		open(page = 'info') {
			element.hidden = false;
			show(page);
			back.focus();
		},
		close() {
			element.hidden = true;
		},
		get isOpen() {
			return !element.hidden;
		},
	};
}
