import { shortcutHint } from './binding-labels';
import { emit } from './events';
import { setComboValue, type ComboInput } from './ribbon-combo';
import { localizeElement, normalizeEditorLocale } from './localization';
import type { RibbonAction } from './ribbon-action';
import { buildHomePanel } from './ribbon-home';
import { attachRibbonBehavior } from './ribbon-behavior';
import { attachRibbonOverflow, fitPanel, refitRibbon } from './ribbon-overflow';
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
	for (const name of ['Home', 'Insert', 'Layout', 'References', 'Review', 'View', 'Table']) {
		const id = `dve-tab-${name.toLowerCase()}`;
		const tab = document.createElement('button');
		tab.type = 'button';
		tab.id = id;
		tab.dataset.tabKey = `tab.${name.toLowerCase()}`;
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
		// Table tools are contextual: the tab appears only while the selection is in a table.
		if (name === 'Table') {
			tab.hidden = true;
			tab.dataset.contextual = '';
		}
		tab.addEventListener('click', () => {
			tabs.querySelectorAll('[role=tab]').forEach((item) => {
				item.setAttribute('aria-selected', String(item === tab));
				(item as HTMLButtonElement).tabIndex = item === tab ? 0 : -1;
			});
			panels.forEach((item, key) => {
				item.hidden = key !== name;
			});
			// Fold groups now, so the new tab never shows a frame of clipped controls.
			fitPanel(panel);
		});
		tab.addEventListener('keydown', (event) => {
			const tabsList = [...tabs.querySelectorAll<HTMLButtonElement>('[role=tab]')].filter(
				(item) => !item.hidden && !item.hasAttribute('data-dve-hidden'),
			);
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
			const target = tabsList[next];
			if (!target) return;
			target.focus();
			target.click();
		});
		tabs.append(tab);
		panels.set(name, panel);
	}
	buildHomePanel(panels);
	buildOtherPanels(panels);
	for (const panel of panels.values()) root.append(panel);
	setComboValue(root.querySelector<ComboInput>('[aria-label="Font family"]')!, 'Calibri');
	setComboValue(root.querySelector<ComboInput>('[aria-label="Font size"]')!, '11');
	root.querySelector<HTMLSelectElement>('[aria-label="Zoom"]')!.value = '100';
	root.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!.value = 'draft';
	root.prepend(tabs);
	attachRibbonBehavior(root, tabs);
	attachRibbonOverflow(root);
	root.addEventListener('click', (event) => {
		const target = (event.target as HTMLElement).closest<HTMLButtonElement>('button[data-action]');
		if (!target) return;
		emit(target, 'ribbon-action', JSON.parse(target.dataset.action!) as RibbonAction);
	});
	for (const type of ['pointerover', 'focusin', 'change'])
		root.addEventListener(type, (event) => {
			if (event.target instanceof HTMLSelectElement) syncSelectTitles(event.target.parentElement!);
		});
	setRibbonLocale(root, locale);
	return root;
}

/** Gives every select the tooltip of its selected label so an ellipsized value stays readable. */
export function syncSelectTitles(root: ParentNode): void {
	for (const select of root.querySelectorAll('select')) {
		const text = select.selectedOptions[0]?.textContent?.trim() ?? '';
		if (text && select.title !== text) select.title = text;
	}
}

/** Gives every text button its full label as a tooltip, for when a narrow slot truncates it. */
function syncButtonTitles(root: ParentNode): void {
	for (const button of root.querySelectorAll('button')) {
		const label = button.getAttribute('aria-label');
		if (!label) continue;
		const hint = shortcutHint(button.dataset.localearialabel ?? label);
		const title = hint ? `${label} (${hint})` : label;
		if ((button.textContent?.trim() || hint) && button.title !== title) button.title = title;
	}
}

export function setRibbonLocale(root: HTMLElement, value: string): void {
	const locale = normalizeEditorLocale(value);
	root.dataset.editorLocale = locale;
	localizeElement(root, locale);
	syncSelectTitles(root);
	syncButtonTitles(root);
	refitRibbon(root);
}
