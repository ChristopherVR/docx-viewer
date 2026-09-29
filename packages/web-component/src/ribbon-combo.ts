import { emit } from './events';
import type { RibbonAction } from './ribbon-action';
import { openListPopover } from './ribbon-popover';
import { ribbonIcon } from './ribbon-icons';

export const FONT_FAMILIES = [
	'Arial',
	'Arial Black',
	'Calibri',
	'Cambria',
	'Candara',
	'Comic Sans MS',
	'Consolas',
	'Constantia',
	'Corbel',
	'Courier New',
	'Garamond',
	'Georgia',
	'Impact',
	'Lucida Console',
	'Palatino Linotype',
	'Segoe UI',
	'Tahoma',
	'Times New Roman',
	'Trebuchet MS',
	'Verdana',
];

/** Word's font size ladder, offered in the list; any size from 1 to 1638 pt can be typed. */
export const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

/** The size a typed value means, or null when it is not a valid point size. */
export function parseFontSize(text: string): number | null {
	const value = Number(text.trim().replace(',', '.'));
	if (!Number.isFinite(value) || value < 1 || value > 1638) return null;
	return Math.round(value * 2) / 2;
}

/** Family names are kept to what Word stores in `w:rFonts`: a single plain name. */
export function parseFontFamily(text: string): string | null {
	const name = text.trim().replace(/\s+/g, ' ');
	return name && name.length <= 64 && !/[;{}"<>\\]/.test(name) ? name : null;
}

export interface ComboInput extends HTMLInputElement {
	/** Entries offered by the drop-down, in order; the current value is added when it is missing. */
	comboItems: string[];
}

/**
 * Word's editable combo box: type a value and press Enter (or leave the box) to apply it, or pick
 * from the list. The input carries the control's name and `value`; the caret button opens the list.
 * `preview` draws each list entry in its own face (font names).
 */
export function comboBox(
	label: string,
	items: readonly string[],
	parse: (text: string) => string | null,
	action: (value: string) => RibbonAction,
	options: { preview?: boolean; className?: string } = {},
) {
	const wrap = document.createElement('div');
	wrap.className = `ribbon-combo ${options.className ?? ''}`.trim();
	const input = document.createElement('input') as ComboInput;
	input.type = 'text';
	input.setAttribute('role', 'combobox');
	input.setAttribute('aria-label', label);
	input.setAttribute('aria-autocomplete', 'list');
	input.setAttribute('aria-haspopup', 'listbox');
	input.autocomplete = 'off';
	input.spellcheck = false;
	input.comboItems = [...items];
	let committed = '';
	const commit = (raw: string) => {
		const value = parse(raw);
		if (value === null) input.value = committed;
		else {
			committed = value;
			input.value = value;
			emit(input, 'ribbon-action', action(value));
		}
	};
	input.addEventListener('focus', () => input.select());
	input.addEventListener('keydown', (event) => {
		if (event.key === 'Enter') {
			event.preventDefault();
			commit(input.value);
		} else if (event.key === 'Escape') input.value = committed;
		else if (event.key === 'ArrowDown') {
			event.preventDefault();
			toggle();
		}
	});
	input.addEventListener('change', () => commit(input.value));
	// Keep what the editor last showed, so an abandoned edit reverts instead of applying.
	input.addEventListener('dve-sync', () => (committed = input.value));
	const caret = document.createElement('button');
	caret.type = 'button';
	caret.className = 'ribbon-caret ribbon-combo-caret';
	caret.dataset.splitCaret = '';
	caret.title = label;
	caret.tabIndex = -1;
	caret.setAttribute('aria-hidden', 'true');
	caret.append(ribbonIcon('caret', 12));
	caret.addEventListener('mousedown', (event) => event.preventDefault());
	const toggle = () =>
		openListPopover(wrap, input.comboItems, input.value, options.preview ?? false, (value) => {
			input.focus();
			commit(value);
		});
	caret.addEventListener('click', toggle);
	wrap.append(input, caret);
	return wrap;
}

/** Shows `value` in a combo (or blank when the selection mixes values) and lists it if it is new. */
export function setComboValue(input: ComboInput, value: string | null): void {
	if (value !== null && !input.comboItems.includes(value))
		input.comboItems = [...input.comboItems, value].sort((a, b) =>
			a.localeCompare(b, undefined, { numeric: true }),
		);
	if (document.activeElement !== input) {
		input.value = value ?? '';
		input.dispatchEvent(new Event('dve-sync'));
	}
}
