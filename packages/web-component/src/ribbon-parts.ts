import { emit } from './events';
import type { RibbonAction } from './ribbon-action';
import { ribbonIcon, type RibbonIcon } from './ribbon-icons';
import { openSwatchPopover } from './ribbon-popover';
import { swatchColor } from './ribbon-colors';

/** Building blocks for ribbon tabs: buttons, selects, labelled groups and rows. */
export const button = (label: string, text: string, action: RibbonAction, className = '') => {
	const el = document.createElement('button');
	el.type = 'button';
	el.textContent = text;
	el.setAttribute('aria-label', label);
	el.addEventListener('mousedown', (event) => event.preventDefault());
	el.dataset.action = JSON.stringify(action);
	if (className) el.className = className;
	return el;
};
export const select = (
	label: string,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction,
) => {
	const el = document.createElement('select');
	el.setAttribute('aria-label', label);
	const stableControlClass: Record<string, string> = {
		'Font family': 'font-family-select',
		'Font size': 'font-size-select',
		'Spacing after': 'spacing-select',
		'Spacing before': 'spacing-select',
		'Line spacing': 'line-spacing-select',
	};
	if (stableControlClass[label]) el.classList.add(stableControlClass[label]);
	for (const [value, text] of values) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = text;
		el.append(option);
	}
	el.dataset.action = 'select';
	el.addEventListener('change', () => emit(el, 'ribbon-action', action(el.value)));
	return el;
};
/**
 * Word's dropdown command: an icon over a caption with a caret. The real `<select>` stays in the
 * DOM, transparent and covering the whole control, so the browser's native list opens on click,
 * assistive technology sees a combobox, and the editor keeps reading and writing `select.value`.
 * `compact` drops the caption, leaving an icon and caret (Word's Line and Paragraph Spacing).
 */
export const menuSelect = (
	label: string,
	icon: RibbonIcon,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction,
	options: { compact?: boolean; momentary?: boolean } = {},
) => menuAround(select(label, values, action), label, icon, options);

/** Wraps an existing select in the dropdown-command look (see `menuSelect`). */
export const menuAround = (
	control: HTMLSelectElement,
	label: string,
	icon: RibbonIcon,
	options: { compact?: boolean; momentary?: boolean } = {},
) => {
	const wrap = document.createElement('div');
	wrap.className = options.compact ? 'ribbon-menu ribbon-menu-compact' : 'ribbon-menu';
	control.title = label;
	control.classList.add('ribbon-menu-input');
	wrap.append(ribbonIcon(icon, options.compact ? 16 : 28));
	if (!options.compact) {
		const caption = document.createElement('span');
		caption.textContent = label;
		wrap.append(caption);
	}
	wrap.append(ribbonIcon('caret', 12), control);
	// A command menu (Change Case) does not keep a value, so choosing the same item again must fire.
	if (options.momentary) {
		control.selectedIndex = -1;
		control.addEventListener('change', () => (control.selectedIndex = -1));
	}
	return wrap;
};
/** A labelled number field (Word's indent spinners). `label` is both the visible text and the name. */
export const spinner = (
	label: string,
	action: (value: number) => RibbonAction,
	options: { min: number; max: number; step: number },
) => {
	const wrap = document.createElement('label');
	wrap.className = 'ribbon-spinner';
	const text = document.createElement('span');
	text.textContent = label;
	const input = document.createElement('input');
	input.type = 'number';
	input.min = String(options.min);
	input.max = String(options.max);
	input.step = String(options.step);
	input.setAttribute('aria-label', label);
	input.addEventListener('change', () => {
		if (input.value !== '' && input.checkValidity())
			emit(input, 'ribbon-action', action(Number(input.value)));
	});
	wrap.append(text, input);
	return wrap;
};
/** Word's dialog launcher: the small corner button of a group that opens its full dialog. */
export const launcher = (label: string, action: RibbonAction) => {
	const el = button(label, '', action, 'ribbon-launcher');
	el.title = label;
	el.append(ribbonIcon('launcher', 10));
	return el;
};
export const group = (label: string, ...children: HTMLElement[]) => {
	const el = document.createElement('div');
	el.className = 'ribbon-group';
	el.dataset.label = label;
	el.dataset.caption = label;
	el.setAttribute('role', 'group');
	el.setAttribute('aria-label', `${label} controls`);
	el.append(...children);
	return el;
};
export const row = (...children: HTMLElement[]) => {
	const el = document.createElement('div');
	el.className = 'ribbon-row';
	el.append(...children);
	return el;
};

/** An icon button. `large` stacks a caption under a 28px icon, as Word's big commands do. */
export const tool = (
	label: string,
	icon: RibbonIcon,
	action: RibbonAction,
	options: { large?: boolean; inline?: boolean; caption?: string; className?: string } = {},
) => {
	const el = button(label, '', action, options.className);
	el.title = label;
	el.classList.add(
		options.large ? 'ribbon-large' : options.inline ? 'ribbon-inline' : 'ribbon-small',
	);
	el.append(ribbonIcon(icon, options.large ? 28 : 16));
	if (options.large || options.inline) {
		const caption = document.createElement('span');
		caption.textContent = options.caption ?? label;
		el.append(caption);
	}
	return el;
};

/** A stack of controls filling a group's height, such as Paste's Cut/Copy column. */
export const stack = (...children: HTMLElement[]) => {
	const el = document.createElement('div');
	el.className = 'ribbon-stack';
	el.append(...children);
	return el;
};

/**
 * Word's split colour button: the main half applies the last chosen colour (shown as a bar under
 * the icon) and the caret opens a swatch popover. `choices` are `[value, English label]` pairs.
 */
export const colorSplit = (
	label: string,
	icon: RibbonIcon,
	choices: Array<[string, string]>,
	action: (value: string) => RibbonAction,
	initial: string,
) => {
	const wrap = document.createElement('div');
	wrap.className = 'ribbon-split';
	const main = button(label, '', action(initial));
	main.removeAttribute('data-action');
	main.title = label;
	main.classList.add('ribbon-small', 'ribbon-color');
	main.dataset.value = initial;
	main.append(ribbonIcon(icon, 16));
	const setBar = (value: string) => {
		main.dataset.value = value;
		main.style.setProperty('--bar', swatchColor(value));
	};
	setBar(initial);
	const apply = (value: string) => {
		setBar(value);
		emit(main, 'ribbon-action', action(value));
	};
	main.addEventListener('click', () => apply(main.dataset.value ?? initial));
	const caret = document.createElement('button');
	caret.type = 'button';
	caret.className = 'ribbon-caret';
	caret.dataset.splitCaret = '';
	caret.setAttribute('aria-haspopup', 'true');
	caret.title = `${label} options`;
	caret.append(ribbonIcon('caret', 12));
	caret.addEventListener('mousedown', (event) => event.preventDefault());
	caret.addEventListener('click', () => openSwatchPopover(caret, choices, apply));
	wrap.append(main, caret);
	return wrap;
};

/** Records `value` as the colour a split button applies next and draws it in the bar. */
export function setColorControl(main: HTMLElement, value: string): void {
	main.dataset.value = value;
	main.style.setProperty('--bar', swatchColor(value));
}
