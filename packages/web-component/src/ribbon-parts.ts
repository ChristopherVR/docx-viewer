import type { RibbonAction } from './ribbon-action';

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
	el.addEventListener('change', () =>
		el.dispatchEvent(
			new CustomEvent('ribbon-action', { bubbles: true, composed: true, detail: action(el.value) }),
		),
	);
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
