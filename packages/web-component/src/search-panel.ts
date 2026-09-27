import type { EditorView } from 'prosemirror-view';
import { SearchController, type SearchStatus } from './search-controller';
import {
	localizeElement,
	normalizeEditorLocale,
	translate,
	type EditorLocale,
} from './localization';

export interface SearchPanelOptions {
	getView: () => EditorView | undefined;
	onClose?: () => void;
}

export interface SearchPanelHandle {
	element: HTMLElement;
	open(): void;
	close(): void;
	refresh(): void;
	setLocale(locale: string): void;
	readonly isOpen: boolean;
}

const styleText = `
	.dve-search-panel{display:flex;align-items:end;flex-wrap:wrap;gap:8px;padding:9px 12px;border-bottom:1px solid var(--line,#ddd);background:var(--surface,#fff);color:var(--ink,#222);font:12px/1.35 'Segoe UI',Arial,sans-serif}
	.dve-search-panel[hidden]{display:none}
	.dve-search-panel label{display:flex;flex-direction:column;gap:3px;color:var(--muted,#666)}
	.dve-search-panel input[type=search],.dve-search-panel input[type=text]{box-sizing:border-box;width:clamp(130px,20vw,220px);height:29px;padding:4px 7px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit}
	.dve-search-panel .case{flex-direction:row;align-items:center;align-self:center;gap:5px;white-space:nowrap}
	.dve-search-panel button{height:29px;padding:0 9px;border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit;cursor:pointer}
	.dve-search-panel button:hover:not(:disabled){border-color:var(--blue,#1769aa)}
	.dve-search-panel button:disabled{opacity:.55;cursor:default}
	.dve-search-panel [role=status]{min-width:100px;color:var(--muted,#666)}
`;

export function createSearchPanel(
	options: SearchPanelOptions,
	initialLocale = 'en',
): SearchPanelHandle {
	let locale: EditorLocale = normalizeEditorLocale(initialLocale);
	const controller = new SearchController(options.getView);
	const panel = document.createElement('section');
	panel.className = 'dve-search-panel';
	panel.setAttribute('role', 'region');
	panel.setAttribute('aria-label', 'Find and replace');
	panel.hidden = true;
	const style = document.createElement('style');
	style.textContent = styleText;
	const findLabel = document.createElement('label');
	findLabel.textContent = 'Find text';
	const findInput = document.createElement('input');
	findInput.type = 'search';
	findInput.setAttribute('aria-label', 'Find text');
	findLabel.append(findInput);
	const replaceLabel = document.createElement('label');
	replaceLabel.textContent = 'Replace with';
	const replaceInput = document.createElement('input');
	replaceInput.type = 'text';
	replaceInput.setAttribute('aria-label', 'Replace with');
	replaceLabel.append(replaceInput);
	const caseLabel = document.createElement('label');
	caseLabel.className = 'case';
	const caseInput = document.createElement('input');
	caseInput.type = 'checkbox';
	caseInput.setAttribute('aria-label', 'Match case');
	caseLabel.append(caseInput, document.createTextNode('Match case'));
	const button = (label: string) => {
		const el = document.createElement('button');
		el.type = 'button';
		el.textContent = label;
		el.setAttribute('aria-label', label);
		return el;
	};
	const previous = button('Find previous');
	const next = button('Find next');
	const replace = button('Replace');
	const replaceAll = button('Replace all');
	const close = button('Close search');
	const result = document.createElement('span');
	result.setAttribute('role', 'status');
	result.setAttribute('aria-label', 'Search results');
	result.setAttribute('aria-live', 'polite');

	const renderStatus = (status: SearchStatus) => {
		result.textContent = !status.query
			? translate(locale, 'Enter text to search')
			: !status.count
				? translate(locale, 'No matches')
				: status.active
					? `${status.active} ${translate(locale, 'of')} ${status.count} ${translate(locale, status.count === 1 ? 'matches one' : 'matches many')}`
					: `${status.count} ${translate(locale, status.count === 1 ? 'matches one' : 'matches many')}`;
		replace.disabled = status.readOnly || !status.count;
		replaceAll.disabled = status.readOnly || !status.count;
		previous.disabled = next.disabled = !status.count;
	};
	const refreshQuery = () => renderStatus(controller.search(findInput.value, caseInput.checked));
	findInput.addEventListener('input', refreshQuery);
	caseInput.addEventListener('change', refreshQuery);
	previous.addEventListener('click', () => {
		controller.findPrevious();
		renderStatus(controller.status);
	});
	next.addEventListener('click', () => {
		controller.findNext();
		renderStatus(controller.status);
	});
	replace.addEventListener('click', () => {
		controller.replaceOne(replaceInput.value);
		renderStatus(controller.status);
	});
	replaceAll.addEventListener('click', () => {
		controller.replaceAll(replaceInput.value);
		renderStatus(controller.status);
	});
	findInput.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			if (event.shiftKey) previous.click();
			else next.click();
		}
	});
	panel.addEventListener('keydown', (event) => {
		if (event.key === 'Escape' && !panel.hidden) {
			event.preventDefault();
			closePanel();
		}
	});

	let isOpen = false;
	const closePanel = () => {
		isOpen = false;
		panel.hidden = true;
		options.onClose?.();
	};
	close.addEventListener('click', closePanel);
	panel.append(
		style,
		findLabel,
		replaceLabel,
		caseLabel,
		previous,
		next,
		replace,
		replaceAll,
		result,
		close,
	);
	localizeElement(panel, locale);
	renderStatus(controller.status);
	return {
		element: panel,
		get isOpen() {
			return isOpen;
		},
		open() {
			isOpen = true;
			panel.hidden = false;
			this.refresh();
			findInput.focus();
		},
		close: closePanel,
		refresh() {
			controller.search(findInput.value, caseInput.checked);
			renderStatus(controller.status);
		},
		setLocale(value: string) {
			locale = normalizeEditorLocale(value);
			localizeElement(panel, locale);
			renderStatus(controller.status);
		},
	};
}
