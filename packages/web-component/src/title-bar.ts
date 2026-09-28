import { iconButton, icon } from './chrome-icons';
import { translateUiText } from './localization';
import type { FileCommand } from './file-commands';

export type SaveState = 'saved' | 'dirty' | 'saving' | 'saved-local';

export interface TitleBarHandlers {
	fileCommand(command: FileCommand): void;
	history(key: 'undo' | 'redo'): void;
	toggleComments(): void;
	setReadOnly(readOnly: boolean): void;
	/** The ribbon root, searched by the "Tell me" box for runnable commands. */
	ribbon(): HTMLElement | undefined;
}

export interface TitleBar {
	element: HTMLElement;
	setFileName(name: string): void;
	setSaveState(state: SaveState): void;
	setReadOnly(readOnly: boolean): void;
	setCommentsOpen(open: boolean): void;
}

const SAVE_STATE_TEXT: Record<SaveState, string> = {
	saved: 'Saved',
	dirty: 'Unsaved changes',
	saving: 'Saving…',
	'saved-local': 'Saved to this PC',
};

/** Ribbon commands the "Tell me" box can run: every labelled, enabled button or select. */
function ribbonCommands(ribbon: HTMLElement): { label: string; control: HTMLElement }[] {
	const seen = new Set<string>();
	const commands: { label: string; control: HTMLElement }[] = [];
	for (const control of ribbon.querySelectorAll<HTMLElement>(
		'.ribbon-panel button[aria-label]:not([data-dve-hidden]), .ribbon-panel select[aria-label]:not([data-dve-hidden])',
	)) {
		const label = control.getAttribute('aria-label')!;
		if (seen.has(label)) continue;
		seen.add(label);
		commands.push({ label, control });
	}
	return commands;
}

/** Shows the ribbon tab holding `control`, then activates (button) or focuses (select) it. */
function runCommand(ribbon: HTMLElement, control: HTMLElement): void {
	const panel = control.closest<HTMLElement>('.ribbon-panel');
	const tab = panel && ribbon.querySelector<HTMLButtonElement>(`[aria-controls="${panel.id}"]`);
	tab?.click();
	if (control instanceof HTMLSelectElement) control.focus();
	else (control as HTMLButtonElement).click();
}

function createTellMe(handlers: TitleBarHandlers): HTMLElement {
	const box = document.createElement('div');
	box.className = 'dve-tellme';
	box.append(icon('search'));
	const input = document.createElement('input');
	input.type = 'search';
	input.setAttribute('aria-label', 'Tell me what you want to do');
	input.placeholder = 'Tell me what you want to do';
	input.setAttribute('role', 'combobox');
	input.setAttribute('aria-expanded', 'false');
	input.setAttribute('aria-controls', 'dve-tellme-results');
	input.setAttribute('aria-autocomplete', 'list');
	const list = document.createElement('ul');
	list.id = 'dve-tellme-results';
	list.className = 'dve-tellme-results';
	list.setAttribute('role', 'listbox');
	list.hidden = true;
	let matches: { label: string; control: HTMLElement }[] = [];
	let active = 0;
	const close = () => {
		list.hidden = true;
		input.setAttribute('aria-expanded', 'false');
	};
	const run = (index: number) => {
		const ribbon = handlers.ribbon();
		const match = matches[index];
		if (!ribbon || !match) return;
		close();
		input.value = '';
		runCommand(ribbon, match.control);
	};
	const render = () => {
		const query = input.value.trim().toLocaleLowerCase();
		const ribbon = handlers.ribbon();
		matches =
			query && ribbon
				? ribbonCommands(ribbon)
						.filter(({ label, control }) => {
							const disabled = (control as HTMLButtonElement).disabled;
							return !disabled && label.toLocaleLowerCase().includes(query);
						})
						.slice(0, 8)
				: [];
		active = 0;
		list.replaceChildren(
			...matches.map(({ label }, index) => {
				const item = document.createElement('li');
				item.setAttribute('role', 'option');
				item.id = `dve-tellme-option-${index}`;
				item.textContent = label;
				item.setAttribute('aria-selected', String(index === active));
				item.addEventListener('mousedown', (event) => event.preventDefault());
				item.addEventListener('click', () => run(index));
				return item;
			}),
		);
		if (query && !matches.length) {
			const empty = document.createElement('li');
			empty.className = 'dve-tellme-empty';
			empty.textContent = translateUiText(
				box.closest<HTMLElement>('.dve-titlebar') ?? box,
				'No matching commands',
			);
			list.append(empty);
		}
		list.hidden = !query;
		input.setAttribute('aria-expanded', String(!list.hidden));
		input.setAttribute('aria-activedescendant', matches.length ? 'dve-tellme-option-0' : '');
	};
	input.addEventListener('input', render);
	input.addEventListener('blur', close);
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			input.value = '';
			close();
		} else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
			if (!matches.length) return;
			event.preventDefault();
			active = (active + (event.key === 'ArrowDown' ? 1 : matches.length - 1)) % matches.length;
			list
				.querySelectorAll('[role=option]')
				.forEach((item, index) => item.setAttribute('aria-selected', String(index === active)));
			input.setAttribute('aria-activedescendant', `dve-tellme-option-${active}`);
		} else if (event.key === 'Enter') {
			event.preventDefault();
			run(active);
		}
	});
	box.append(input, list);
	return box;
}

/** Word-style title bar: quick access toolbar, document name/state, "Tell me" and view mode. */
export function createTitleBar(handlers: TitleBarHandlers): TitleBar {
	const element = document.createElement('header');
	element.className = 'dve-titlebar';
	const quick = document.createElement('div');
	quick.className = 'dve-quick-access';
	quick.setAttribute('role', 'toolbar');
	quick.setAttribute('aria-label', 'Quick access');
	const badge = document.createElement('span');
	badge.className = 'dve-app-badge';
	badge.textContent = 'W';
	badge.setAttribute('aria-hidden', 'true');
	const save = iconButton('save', 'Save');
	save.dataset.fileCommand = 'save';
	save.addEventListener('click', () => handlers.fileCommand('save'));
	const undo = iconButton('undo', 'Undo');
	undo.addEventListener('click', () => handlers.history('undo'));
	const redo = iconButton('redo', 'Redo');
	redo.addEventListener('click', () => handlers.history('redo'));
	for (const button of [save, undo, redo])
		button.addEventListener('mousedown', (event) => event.preventDefault());
	quick.append(badge, save, undo, redo);

	const title = document.createElement('div');
	title.className = 'dve-document-title';
	const name = document.createElement('span');
	name.className = 'dve-filename';
	const separator = document.createElement('span');
	separator.className = 'dve-title-separator';
	separator.textContent = '•';
	separator.setAttribute('aria-hidden', 'true');
	const state = document.createElement('span');
	state.className = 'dve-save-state';
	state.setAttribute('aria-live', 'polite');
	title.append(name, separator, state);

	const actions = document.createElement('div');
	actions.className = 'dve-title-actions';
	const comments = iconButton('comment', 'Show comments');
	comments.setAttribute('aria-pressed', 'false');
	comments.addEventListener('click', () => handlers.toggleComments());
	const mode = document.createElement('select');
	mode.className = 'dve-mode-select';
	mode.setAttribute('aria-label', 'Editing mode');
	const modeOptions: [string, string][] = [
		['editing', 'Editing'],
		['viewing', 'Viewing'],
	];
	for (const [value, text] of modeOptions) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = text;
		mode.append(option);
	}
	mode.addEventListener('change', () => handlers.setReadOnly(mode.value === 'viewing'));
	actions.append(comments, mode);

	const start = document.createElement('div');
	start.className = 'dve-titlebar-start';
	start.append(quick, title);
	element.append(start, createTellMe(handlers), actions);
	return {
		element,
		setFileName(value) {
			name.textContent = value;
			name.title = value;
		},
		setSaveState(value) {
			state.dataset.state = value;
			state.textContent = translateUiText(element, SAVE_STATE_TEXT[value]);
		},
		setReadOnly(readOnly) {
			mode.value = readOnly ? 'viewing' : 'editing';
			save.disabled = false;
			undo.disabled = readOnly;
			redo.disabled = readOnly;
		},
		setCommentsOpen(open) {
			comments.setAttribute('aria-pressed', String(open));
		},
	};
}
