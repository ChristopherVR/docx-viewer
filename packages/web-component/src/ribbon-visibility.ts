import type { RibbonActionId } from './ribbon-action-ids';

const HIDDEN = 'data-dve-hidden';
const CONTROLS = 'button[aria-label], select[aria-label], input[aria-label]';

/** English label of a ribbon control, stable across locales. */
export function ribbonControlId(control: HTMLElement): string | null {
	return control.dataset.localearialabel ?? control.getAttribute('aria-label');
}

/**
 * Applies `showToolbar` and `hiddenActions` to the ribbon. Groups and tabs whose controls are all
 * hidden are hidden too, and the active tab moves to the first visible one. Unknown ids are ignored.
 */
export function applyRibbonVisibility(
	toolbar: HTMLElement,
	showToolbar: boolean,
	hiddenActions: readonly RibbonActionId[],
): void {
	toolbar.toggleAttribute(HIDDEN, !showToolbar);
	const hidden = new Set<string>(hiddenActions);
	const allHidden = (root: Element) => {
		const controls = [...root.querySelectorAll<HTMLElement>(CONTROLS)];
		return controls.length > 0 && controls.every((control) => control.hasAttribute(HIDDEN));
	};
	for (const control of toolbar.querySelectorAll<HTMLElement>(`.ribbon-panel ${CONTROLS}`))
		control.toggleAttribute(HIDDEN, hidden.has(ribbonControlId(control) ?? ''));
	for (const group of toolbar.querySelectorAll('.ribbon-group'))
		group.toggleAttribute(HIDDEN, allHidden(group));
	const tabs = [...toolbar.querySelectorAll<HTMLButtonElement>('[role="tab"]')];
	for (const tab of tabs) {
		const panel = toolbar.querySelector(`#${tab.getAttribute('aria-controls')}`);
		tab.toggleAttribute(HIDDEN, Boolean(panel) && allHidden(panel!));
	}
	const selected = tabs.find((tab) => tab.getAttribute('aria-selected') === 'true');
	if (selected?.hasAttribute(HIDDEN)) tabs.find((tab) => !tab.hasAttribute(HIDDEN))?.click();
}
