import { localizeElement, normalizeEditorLocale } from './localization';
import type { RibbonAction } from './ribbon-action';
import { buildHomePanel } from './ribbon-home';
import { buildOtherPanels } from './ribbon-tabs';

export type { RibbonAction } from './ribbon-action';

export function createRibbon(locale: string = 'en'): HTMLElement {
	const root = document.createElement('div');
	root.className = 'dve-ribbon';
	root.setAttribute('role', 'toolbar');
	root.setAttribute('aria-label', 'Document formatting');
	const tabs = document.createElement('nav');
	tabs.className = 'ribbon-tabs';
	tabs.setAttribute('role', 'tablist');
	tabs.setAttribute('aria-label', 'Ribbon tabs');
	const panels = new Map<string, HTMLElement>();
	for (const name of ['Home', 'Insert', 'Layout', 'Review', 'View', 'Table']) {
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
	buildHomePanel(panels);
	buildOtherPanels(panels);
	for (const panel of panels.values()) root.append(panel);
	root.querySelector<HTMLSelectElement>('[aria-label="Font family"]')!.value = 'Calibri';
	root.querySelector<HTMLSelectElement>('[aria-label="Font size"]')!.value = '11';
	root.querySelector<HTMLSelectElement>('[aria-label="Font color"]')!.value = '#000000';
	root.querySelector<HTMLSelectElement>('[aria-label="Text highlight"]')!.value = 'none';
	root.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!.value = '100';
	root.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!.value = 'draft';
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
	setRibbonLocale(root, locale);
	return root;
}

export function setRibbonLocale(root: HTMLElement, value: string): void {
	const locale = normalizeEditorLocale(value);
	root.dataset.editorLocale = locale;
	localizeElement(root, locale);
}
