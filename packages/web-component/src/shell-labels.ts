import { translate, type EditorLocale } from './localization';

function setLabel(element: HTMLElement | undefined, value: string): void {
	if (element && element.getAttribute('aria-label') !== value)
		element.setAttribute('aria-label', value);
}

/** Localizes the accessible names of the host element and the page surface. */
export function applyShellLabels(
	host: HTMLElement,
	paper: HTMLElement | undefined,
	locale: EditorLocale,
): void {
	setLabel(host, translate(locale, 'Document editor'));
	setLabel(paper, translate(locale, 'Document page'));
}
