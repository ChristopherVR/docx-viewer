import { children, first, getW, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

/** `w:background` colours are six hex digits; `auto` and theme-only fills are not modeled. */
const HEX = /^[0-9a-fA-F]{6}$/;

/** The page colour of `w:document/w:background`, as uppercase hex without `#`. */
export function parsePageBackground(root: XmlElement | null | undefined): string | undefined {
	const color = getW(first(root, 'background'), 'color');
	return color && HEX.test(color) ? color.toUpperCase() : undefined;
}

/**
 * Sets or clears the page colour. An unchanged colour leaves the element (and any theme
 * attributes on it) as it was; a new one replaces it, ahead of `w:body` as the schema requires.
 */
export function applyPageBackground(doc: XmlDocument, color: string | undefined): void {
	const root = doc.documentElement as XmlElement;
	const existing = first(root, 'background');
	const hex = color?.replace(/^#/, '').toUpperCase() ?? '';
	if (!HEX.test(hex)) {
		if (existing) root.removeChild(existing);
		return;
	}
	if (existing && getW(existing, 'color')?.toUpperCase() === hex) return;
	const element = makeW(doc, 'background');
	element.setAttributeNS(WORD_NS, 'w:color', hex);
	if (existing) root.replaceChild(element, existing);
	else root.insertBefore(element, children(root, 'body')[0] ?? null);
}
