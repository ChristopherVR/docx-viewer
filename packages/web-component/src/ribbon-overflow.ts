import { ribbonIcon, type RibbonIcon } from './ribbon-icons';

/** The icon shown on a collapsed group's button, by the group's English name. */
const GROUP_ICONS: Record<string, RibbonIcon> = {
	Clipboard: 'paste',
	Font: 'changeCase',
	Paragraph: 'alignLeft',
	Styles: 'moreStyles',
	Editing: 'find',
	Tables: 'table',
	Illustrations: 'picture',
	Links: 'link',
	Breaks: 'pageBreak',
	'Header & Footer': 'header',
	Text: 'dateTime',
	Symbols: 'symbol',
	'Page setup': 'margins',
	'Page numbers': 'pageNumber',
	'Section breaks': 'sectionNext',
	Indent: 'indent',
	Proofing: 'spelling',
	Tracking: 'track',
	Changes: 'accept',
	Comments: 'comment',
	Language: 'spelling',
	'Table of contents': 'toc',
	Captions: 'caption',
	Footnotes: 'footnote',
	Show: 'view',
	Zoom: 'zoom',
	Views: 'view',
	Page: 'pageSize',
};

/** Detaches whichever collapsed-group panel is open. */
let closeOpen: (() => void) | undefined;

function overflowButton(group: HTMLElement): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = 'ribbon-overflow-button ribbon-large';
	button.setAttribute('aria-haspopup', 'true');
	const label = group.dataset.label ?? '';
	button.append(ribbonIcon(GROUP_ICONS[label] ?? 'moreStyles', 28));
	const caption = document.createElement('span');
	button.append(caption, ribbonIcon('caret', 12));
	button.addEventListener('mousedown', (event) => event.preventDefault());
	button.addEventListener('click', () => toggleGroup(group, button));
	return button;
}

/** Shows a collapsed group's controls in a panel under its button, moving the live elements. */
function toggleGroup(group: HTMLElement, button: HTMLButtonElement): void {
	const wasOpen = button.getAttribute('aria-expanded') === 'true';
	closeOpen?.();
	if (wasOpen) return;
	const ribbon = group.closest<HTMLElement>('.dve-ribbon');
	if (!ribbon) return;
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover ribbon-overflow-panel';
	const content = document.createElement('div');
	content.className = 'ribbon-group ribbon-overflow-content';
	content.dataset.label = group.dataset.label ?? '';
	content.dataset.caption = group.dataset.caption ?? '';
	content.setAttribute('role', 'group');
	content.setAttribute('aria-label', group.getAttribute('aria-label') ?? '');
	content.style.width = `${group.dataset.natural ?? 200}px`;
	const moved = [...group.children].filter((child) => child !== button);
	content.append(...moved);
	pop.append(content);
	ribbon.append(pop);
	const box = button.getBoundingClientRect();
	pop.style.top = `${box.bottom + 2}px`;
	const limit = window.innerWidth - pop.offsetWidth - 8;
	pop.style.left = `${Math.max(4, Math.min(box.left, limit))}px`;
	button.setAttribute('aria-expanded', 'true');
	const outside = (event: Event) => {
		const path = event.composedPath();
		const inside = path.some(
			(node) => node instanceof HTMLElement && node.classList.contains('ribbon-popover'),
		);
		if (!inside && !path.includes(button)) closeOpen?.();
	};
	const escape = (event: KeyboardEvent) => {
		if (event.key !== 'Escape') return;
		closeOpen?.();
		button.focus();
	};
	// A command ends the dropdown, as in Word; selects and combo boxes keep it open.
	pop.addEventListener('click', (event) => {
		if ((event.target as HTMLElement).closest('button[data-action]'))
			setTimeout(() => closeOpen?.(), 0);
	});
	document.addEventListener('pointerdown', outside, true);
	document.addEventListener('keydown', escape, true);
	closeOpen = () => {
		document.removeEventListener('pointerdown', outside, true);
		document.removeEventListener('keydown', escape, true);
		group.append(...moved);
		pop.remove();
		button.removeAttribute('aria-expanded');
		closeOpen = undefined;
	};
}

/** Puts every group of `panel` back in place, ready to be measured at its natural width. */
function expand(panel: HTMLElement): void {
	for (const group of panel.querySelectorAll<HTMLElement>('.ribbon-group[data-collapsed]'))
		group.removeAttribute('data-collapsed');
}

/**
 * Collapses groups of a panel from the right, each into one dropdown button, until the panel fits
 * its width; the panel scrolls only when even that is not enough. Word does the same when a window
 * narrows.
 */
export function fitPanel(panel: HTMLElement): void {
	if (panel.hidden || !panel.clientWidth) return;
	// Re-measuring while a folded group's panel is open would tear it down under the user's hands.
	if (closeOpen && panel.dataset.fittedWidth === String(panel.clientWidth)) return;
	panel.dataset.fittedWidth = String(panel.clientWidth);
	closeOpen?.();
	expand(panel);
	const groups = [...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')].filter(
		(group) => group.offsetWidth > 0,
	);
	for (const group of groups) group.dataset.natural = String(Math.ceil(group.offsetWidth));
	for (let at = groups.length - 1; at >= 0 && panel.scrollWidth > panel.clientWidth + 1; at--) {
		const group = groups[at]!;
		let button = group.querySelector<HTMLButtonElement>(':scope > .ribbon-overflow-button');
		if (!button) {
			button = overflowButton(group);
			group.append(button);
		}
		const caption = button.querySelector('span')!;
		caption.textContent = group.dataset.caption ?? group.dataset.label ?? '';
		button.setAttribute('aria-label', caption.textContent);
		button.title = caption.textContent;
		group.setAttribute('data-collapsed', '');
	}
}

/** Keeps each panel fitted as the ribbon resizes or a tab is shown. */
export function attachRibbonOverflow(root: HTMLElement): () => void {
	if (typeof ResizeObserver === 'undefined') return () => {};
	const observer = new ResizeObserver((entries) => {
		for (const entry of entries) fitPanel(entry.target as HTMLElement);
	});
	for (const panel of root.querySelectorAll<HTMLElement>('.ribbon-panel')) observer.observe(panel);
	return () => observer.disconnect();
}

/** Re-measures the visible panel, for example after the display language changed label widths. */
export function refitRibbon(root: HTMLElement): void {
	for (const panel of root.querySelectorAll<HTMLElement>('.ribbon-panel')) fitPanel(panel);
}
