import { dialogButton } from './dialog-fields';

/** Accessible Font/Advanced tabs, with roving keyboard focus. */
export function createFontDialogTabs() {
	const list = document.createElement('div');
	list.className = 'dve-font-tabs';
	list.setAttribute('role', 'tablist');
	const tabs = ['Font', 'Advanced'].map((label) => dialogButton(label));
	const panels = tabs.map(() => document.createElement('div'));
	const prefix = `dve-font-${Math.random().toString(36).slice(2)}`;
	const select = (index: number, focus = false) => {
		tabs.forEach((tab, i) => {
			tab.setAttribute('aria-selected', String(i === index));
			tab.tabIndex = i === index ? 0 : -1;
			panels[i]!.hidden = i !== index;
		});
		if (focus) tabs[index]!.focus();
	};
	tabs.forEach((tab, index) => {
		tab.id = `${prefix}-tab-${index}`;
		tab.setAttribute('role', 'tab');
		tab.setAttribute('aria-controls', `${prefix}-panel-${index}`);
		const panel = panels[index]!;
		panel.id = `${prefix}-panel-${index}`;
		panel.className = 'dve-font-tab-panel';
		panel.setAttribute('role', 'tabpanel');
		panel.setAttribute('aria-labelledby', tab.id);
		tab.addEventListener('click', () => select(index));
		tab.addEventListener('keydown', (event) => {
			if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
			event.preventDefault();
			select(event.key === 'Home' ? 0 : event.key === 'End' ? 1 : 1 - index, true);
		});
	});
	list.append(...tabs);
	select(0);
	return {
		list,
		basic: panels[0]!,
		advanced: panels[1]!,
		reset: () => select(0),
		showAdvanced: () => select(1),
	};
}
