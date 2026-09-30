import { ribbonIcon } from './ribbon-icons';
import { showCommandTips } from './ribbon-keytips';

/** Word's tab KeyTips (Alt, then a letter). File is handled by the chrome's own tab. */
const KEY_TIPS: Record<string, string> = {
	'dve-tab-home': 'H',
	'dve-tab-insert': 'N',
	'dve-tab-layout': 'P',
	'dve-tab-references': 'S',
	'dve-tab-review': 'R',
	'dve-tab-view': 'W',
	'dve-tab-table': 'T',
};

/**
 * Ribbon collapse and KeyTips.
 *
 * Collapse: double-clicking the selected tab, Ctrl+F1 or the collapse button hides the panels; a
 * click on a tab then shows its panel over the document until Escape, a command, or a click
 * elsewhere. KeyTips: when the ribbon is focused with Alt or F10 (the host fires `dve-keytips`),
 * tab badges appear and the matching letter opens that tab (`F` opens File); the tab's commands
 * then show their own tips (see ribbon-keytips.ts).
 */
export function attachRibbonBehavior(root: HTMLElement, tabs: HTMLElement): void {
	const collapse = document.createElement('button');
	collapse.type = 'button';
	collapse.className = 'ribbon-collapse';
	collapse.setAttribute('aria-label', 'Collapse the ribbon');
	collapse.title = 'Collapse the ribbon';
	collapse.setAttribute('aria-pressed', 'false');
	collapse.append(ribbonIcon('previous', 14));
	tabs.append(collapse);

	const outside = (event: Event) => {
		if (!event.composedPath().includes(root)) peek(false);
	};
	/** Shows or hides the overlaid panel of a collapsed ribbon; listens outside only while shown. */
	const peek = (on: boolean) => {
		root.toggleAttribute('data-peek', on);
		if (on) document.addEventListener('pointerdown', outside, true);
		else document.removeEventListener('pointerdown', outside, true);
	};
	const setCollapsed = (value: boolean) => {
		root.toggleAttribute('data-collapsed', value);
		peek(false);
		collapse.setAttribute('aria-pressed', String(value));
	};
	const toggle = () => setCollapsed(!root.hasAttribute('data-collapsed'));
	collapse.addEventListener('click', toggle);
	root.addEventListener('keydown', (event) => {
		if (event.key === 'F1' && event.ctrlKey) {
			event.preventDefault();
			toggle();
		} else if (event.key === 'Escape' && root.hasAttribute('data-peek')) peek(false);
	});

	for (const tab of tabs.querySelectorAll<HTMLElement>('[role="tab"]')) {
		tab.addEventListener('dblclick', toggle);
		tab.addEventListener('click', () => {
			if (root.hasAttribute('data-collapsed')) peek(true);
		});
	}
	root.addEventListener('ribbon-action', () => peek(false));

	let hideTips: (() => void) | undefined;
	root.addEventListener('dve-keytips', () => {
		hideTips?.();
		const badges = [...tabs.querySelectorAll<HTMLElement>('[role="tab"], .dve-file-tab')].flatMap(
			(tab) => {
				const key = tab.classList.contains('dve-file-tab') ? 'F' : KEY_TIPS[tab.id];
				if (!key || tab.hidden || tab.hasAttribute('data-dve-hidden')) return [];
				const badge = document.createElement('span');
				badge.className = 'dve-keytip';
				badge.textContent = key;
				badge.setAttribute('aria-hidden', 'true');
				tab.append(badge);
				return [{ badge, key, tab }];
			},
		);
		const stop = () => {
			for (const { badge } of badges) badge.remove();
			root.removeEventListener('keydown', onKey, true);
			root.removeEventListener('pointerdown', stop, true);
			root.removeEventListener('focusout', onFocusOut);
			hideTips = undefined;
		};
		const onKey = (event: KeyboardEvent) => {
			if (event.key === 'Alt' || event.key === 'Shift') return;
			const match = badges.find(({ key }) => key === event.key.toUpperCase());
			stop();
			if (!match || event.ctrlKey || event.metaKey) return;
			event.preventDefault();
			event.stopPropagation();
			match.tab.click();
			match.tab.focus();
			// Word's second level: the tab's commands get their own tips.
			const panel = root.querySelector<HTMLElement>(`#${match.tab.getAttribute('aria-controls')}`);
			if (panel) showCommandTips(root, panel, () => match.tab.focus());
		};
		const onFocusOut = (event: FocusEvent) => {
			if (!event.relatedTarget || !root.contains(event.relatedTarget as Node)) stop();
		};
		root.addEventListener('keydown', onKey, true);
		root.addEventListener('pointerdown', stop, true);
		root.addEventListener('focusout', onFocusOut);
		hideTips = stop;
	});
}
