import type { TableCommand } from './table-commands';
import { lineSpacingOptions } from './line-spacing';

export type RibbonAction =
	| {
			type: 'format';
			key: 'bold' | 'italic' | 'underline' | 'strike' | 'superscript' | 'subscript';
	  }
	| { type: 'history'; key: 'undo' | 'redo' }
	| { type: 'align'; value: 'left' | 'center' | 'right' | 'justify' }
	| { type: 'font'; key: 'family' | 'size' | 'color' | 'highlight'; value: string }
	| { type: 'clear' | 'table' }
	| { type: 'tableEdit'; key: TableCommand }
	| { type: 'page'; key: 'margin' | 'orientation'; value: string }
	| {
			type: 'paragraph';
			key: 'indent' | 'spacingBefore' | 'spacingAfter' | 'lineSpacing';
			value: string;
	  }
	| { type: 'zoom'; value: number };

const button = (label: string, text: string, action: RibbonAction, className = '') => {
	const el = document.createElement('button');
	el.type = 'button';
	el.textContent = text;
	el.setAttribute('aria-label', label);
	el.addEventListener('mousedown', (event) => event.preventDefault());
	el.dataset.action = JSON.stringify(action);
	if (className) el.className = className;
	return el;
};
const select = (
	label: string,
	values: Array<[string, string]>,
	action: (value: string) => RibbonAction,
) => {
	const el = document.createElement('select');
	el.setAttribute('aria-label', label);
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
const group = (...children: HTMLElement[]) => {
	const el = document.createElement('div');
	el.className = 'ribbon-group';
	el.append(...children);
	return el;
};

export function createRibbon(): HTMLElement {
	const root = document.createElement('div');
	root.className = 'dve-ribbon';
	root.setAttribute('role', 'toolbar');
	root.setAttribute('aria-label', 'Document formatting');
	const tabs = document.createElement('nav');
	tabs.className = 'ribbon-tabs';
	tabs.setAttribute('role', 'tablist');
	tabs.setAttribute('aria-label', 'Ribbon tabs');
	const panels = new Map<string, HTMLElement>();
	for (const name of ['Home', 'Insert', 'Layout', 'View', 'Table']) {
		const id = `dve-tab-${name.toLowerCase()}`;
		const tab = document.createElement('button');
		tab.type = 'button';
		tab.id = id;
		tab.textContent = name;
		tab.setAttribute('role', 'tab');
		tab.setAttribute('aria-selected', String(name === 'Home'));
		tab.setAttribute('aria-controls', `dve-panel-${name.toLowerCase()}`);
		tab.tabIndex = name === 'Home' ? 0 : -1;
		const panel = document.createElement('div');
		panel.className = 'ribbon-panel';
		panel.dataset.panel = name;
		panel.id = `dve-panel-${name.toLowerCase()}`;
		panel.setAttribute('role', 'tabpanel');
		panel.setAttribute('aria-labelledby', id);
		panel.tabIndex = 0;
		panel.hidden = name !== 'Home';
		tab.addEventListener('click', () => {
			tabs.querySelectorAll('[role=tab]').forEach((item) => {
				item.setAttribute('aria-selected', String(item === tab));
				(item as HTMLButtonElement).tabIndex = item === tab ? 0 : -1;
			});
			panels.forEach((item, key) => {
				item.hidden = key !== name;
			});
		});
		tab.addEventListener('keydown', (event) => {
			const tabsList = [...tabs.querySelectorAll<HTMLButtonElement>('[role=tab]')];
			const current = tabsList.indexOf(tab);
			const next =
				event.key === 'ArrowRight'
					? (current + 1) % tabsList.length
					: event.key === 'ArrowLeft'
						? (current + tabsList.length - 1) % tabsList.length
						: event.key === 'Home'
							? 0
							: event.key === 'End'
								? tabsList.length - 1
								: -1;
			if (next < 0) return;
			event.preventDefault();
			tabsList[next].focus();
			tabsList[next].click();
		});
		tabs.append(tab);
		panels.set(name, panel);
	}
	const home = panels.get('Home')!;
	home.append(
		group(
			select(
				'Font family',
				[
					['Arial', 'Arial'],
					['Calibri', 'Calibri'],
					['Georgia', 'Georgia'],
					['Times New Roman', 'Times New Roman'],
					['Verdana', 'Verdana'],
				],
				(value) => ({ type: 'font', key: 'family', value }),
			),
			select(
				'Font size',
				['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'].map(
					(v) => [v, v],
				),
				(value) => ({ type: 'font', key: 'size', value }),
			),
			button('Bold', 'B', { type: 'format', key: 'bold' }, 'tool-bold'),
			button('Italic', 'I', { type: 'format', key: 'italic' }, 'tool-italic'),
			button('Underline', 'U', { type: 'format', key: 'underline' }, 'tool-underline'),
			button('Strikethrough', 'S̶', { type: 'format', key: 'strike' }, 'tool-strike'),
			button('Superscript', 'x²', { type: 'format', key: 'superscript' }),
			button('Subscript', 'x₂', { type: 'format', key: 'subscript' }),
			select(
				'Font color',
				[
					['#000000', 'Black'],
					['#c00000', 'Red'],
					['#e36c09', 'Orange'],
					['#ffc000', 'Gold'],
					['#70ad47', 'Green'],
					['#0070c0', 'Blue'],
					['#7030a0', 'Purple'],
				],
				(value) => ({ type: 'font', key: 'color', value }),
			),
			select(
				'Text highlight',
				[
					['none', 'No highlight'],
					['yellow', 'Yellow'],
					['green', 'Green'],
					['cyan', 'Cyan'],
					['magenta', 'Magenta'],
					['blue', 'Blue'],
					['red', 'Red'],
					['darkBlue', 'Dark blue'],
					['darkCyan', 'Dark cyan'],
					['darkGreen', 'Dark green'],
					['darkMagenta', 'Dark magenta'],
					['darkRed', 'Dark red'],
					['darkYellow', 'Dark yellow'],
					['darkGray', 'Dark gray'],
					['lightGray', 'Light gray'],
					['black', 'Black'],
					['white', 'White'],
				],
				(value) => ({ type: 'font', key: 'highlight', value }),
			),
			button('Clear formatting', 'Clear', { type: 'clear' }),
		),
	);
	panels
		.get('Table')!
		.append(
			group(
				button('Insert row above', '↑ Row', { type: 'tableEdit', key: 'rowBefore' }),
				button('Insert row below', '↓ Row', { type: 'tableEdit', key: 'rowAfter' }),
				button('Delete row', '− Row', { type: 'tableEdit', key: 'deleteRow' }),
				button('Insert column left', '← Column', { type: 'tableEdit', key: 'columnBefore' }),
				button('Insert column right', '→ Column', { type: 'tableEdit', key: 'columnAfter' }),
				button('Delete column', '− Column', { type: 'tableEdit', key: 'deleteColumn' }),
				button('Delete table', 'Delete table', { type: 'tableEdit', key: 'deleteTable' }),
			),
		);
	home.append(
		group(
			button('Decrease indent', '⇤', { type: 'paragraph', key: 'indent', value: 'decrease' }),
			button('Increase indent', '⇥', { type: 'paragraph', key: 'indent', value: 'increase' }),
			select(
				'Spacing after',
				[
					['inherit', 'After: style default'],
					['0', 'After: no paragraph space'],
					['120', 'After: 6 pt'],
					['240', 'After: 12 pt'],
					['360', 'After: 18 pt'],
				],
				(value) => ({ type: 'paragraph', key: 'spacingAfter', value }),
			),
			select(
				'Spacing before',
				[
					['inherit', 'Before: style default'],
					['0', 'Before: no paragraph space'],
					['120', 'Before: 6 pt'],
					['240', 'Before: 12 pt'],
					['360', 'Before: 18 pt'],
				],
				(value) => ({ type: 'paragraph', key: 'spacingBefore', value }),
			),
			select('Line spacing', lineSpacingOptions, (value) => ({
				type: 'paragraph',
				key: 'lineSpacing',
				value,
			})),
		),
	);
	home.append(
		group(
			...(
				[
					['left', 'Align left'],
					['center', 'Align center'],
					['right', 'Align right'],
					['justify', 'Justify'],
				] as const
			).map(([value, label]) =>
				button(label, value === 'center' ? '≣' : value === 'justify' ? '☰' : '≡', {
					type: 'align',
					value,
				}),
			),
			button('Undo', '↶', { type: 'history', key: 'undo' }),
			button('Redo', '↷', { type: 'history', key: 'redo' }),
		),
	);
	panels.get('Insert')!.append(group(button('Insert table', '▦ Table', { type: 'table' })));
	panels.get('Layout')!.append(
		group(
			select(
				'Margins',
				[
					['normal', 'Normal'],
					['narrow', 'Narrow'],
					['wide', 'Wide'],
				],
				(value) => ({ type: 'page', key: 'margin', value }),
			),
			select(
				'Orientation',
				[
					['portrait', 'Portrait'],
					['landscape', 'Landscape'],
				],
				(value) => ({ type: 'page', key: 'orientation', value }),
			),
		),
	);
	panels.get('View')!.append(
		group(
			select(
				'Zoom',
				[
					['50', '50%'],
					['75', '75%'],
					['90', '90%'],
					['100', '100%'],
					['125', '125%'],
					['150', '150%'],
				],
				(value) => ({ type: 'zoom', value: Number(value) }),
			),
		),
	);
	for (const panel of panels.values()) root.append(panel);
	root.querySelector<HTMLSelectElement>('[aria-label="Font family"]')!.value = 'Calibri';
	root.querySelector<HTMLSelectElement>('[aria-label="Font size"]')!.value = '11';
	root.querySelector<HTMLSelectElement>('[aria-label="Font color"]')!.value = '#000000';
	root.querySelector<HTMLSelectElement>('[aria-label="Text highlight"]')!.value = 'none';
	root.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!.value = '100';
	root.prepend(tabs);
	root.addEventListener('click', (event) => {
		const target = (event.target as HTMLElement).closest<HTMLButtonElement>('button[data-action]');
		if (!target) return;
		target.dispatchEvent(
			new CustomEvent('ribbon-action', {
				bubbles: true,
				composed: true,
				detail: JSON.parse(target.dataset.action!) as RibbonAction,
			}),
		);
	});
	return root;
}
