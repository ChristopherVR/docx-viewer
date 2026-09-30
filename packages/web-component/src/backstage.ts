import { icon, type ChromeIcon } from './chrome-icons';
import type { FileCommand } from './file-commands';
import { translateUiText } from './localization';
import {
	renderExport,
	renderInfo,
	renderNew,
	renderOpen,
	renderOptions,
	renderCustomize,
	renderPrint,
	renderSaveAs,
	type BackstageHandlers,
	type PageContext,
} from './backstage-pages';

export type { BackstageHandlers } from './backstage-pages';
export type BackstagePage =
	| 'info'
	| 'new'
	| 'open'
	| 'saveAs'
	| 'print'
	| 'export'
	| 'options'
	| 'customize';

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

const PAGES: Array<[BackstagePage, ChromeIcon, string]> = [
	['info', 'info', 'Info'],
	['new', 'file', 'New'],
	['open', 'folder', 'Open'],
	['saveAs', 'copy', 'Save As'],
	['print', 'print', 'Print'],
	['export', 'file', 'Export'],
	['options', 'settings', 'Options'],
	['customize', 'settings', 'Customize Ribbon'],
];
const RENDERERS: Record<BackstagePage, (context: PageContext) => void> = {
	info: renderInfo,
	new: renderNew,
	open: renderOpen,
	saveAs: renderSaveAs,
	print: renderPrint,
	export: renderExport,
	options: renderOptions,
	customize: renderCustomize,
};

/**
 * Word's File view. Left: Back, Info, New, Open, Save and its siblings (Save As, Print, Export)
 * and Options; right: the selected page. Info shows properties and compatibility notes, Save As
 * takes a file name, Export offers PDF (through printing), DOCX and plain text, and Options sets
 * the editor's language, theme and review author.
 */
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
	const buttons = new Map<BackstagePage, HTMLButtonElement>();
	let current: BackstagePage = 'info';
	// A changed language or theme re-renders the open page so its text follows.
	const pageHandlers: BackstageHandlers = {
		...handlers,
		setOption(key, value) {
			handlers.setOption(key, value);
			queueMicrotask(() => show(current));
		},
	};
	const show = (page: BackstagePage) => {
		current = page;
		for (const [key, button] of buttons) button.setAttribute('aria-current', String(key === page));
		RENDERERS[page]({ handlers: pageHandlers, t, content });
	};
	const separator = () => {
		const line = document.createElement('div');
		line.className = 'dve-backstage-separator';
		line.setAttribute('role', 'separator');
		return line;
	};
	const save = navButton('save', 'Save');
	save.addEventListener('click', () => {
		handlers.close();
		handlers.fileCommand('save' satisfies FileCommand);
	});
	nav.append(back);
	for (const [page, iconName, label] of PAGES) {
		const button = navButton(iconName, label);
		button.addEventListener('click', () => show(page));
		buttons.set(page, button);
		nav.append(button);
		// Save sits with Save As, Print and Export as in Word's list of file actions.
		if (page === 'open') nav.append(separator(), save);
		if (page === 'export') nav.append(separator());
	}
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
