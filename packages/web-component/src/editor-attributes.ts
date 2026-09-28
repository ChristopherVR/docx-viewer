import { normalizeEditorLocale } from './localization';
import { normalizeThemeMode, type EditorThemeMode } from './theme';

/**
 * Attribute <-> property mapping for `<docx-editor>`. Properties stay the source of truth: an
 * attribute change is applied through the property setter only when it would change the value, and
 * a property change is written back to its attribute only when the attribute would differ. That
 * makes each direction idempotent, so the two cannot ping-pong.
 */
export const DOCX_EDITOR_ATTRIBUTES = [
	'locale',
	'read-only',
	'file-name',
	'review-author',
	'theme',
] as const;
export type DocxEditorAttribute = (typeof DOCX_EDITOR_ATTRIBUTES)[number];

export const DEFAULT_FILE_NAME = 'Document1.docx';
export const DEFAULT_REVIEW_AUTHOR = 'Author';

/** The element properties the attributes drive. */
export interface AttributeProperties {
	locale: string;
	readOnly: boolean;
	fileName: string;
	reviewAuthor: string;
	theme: EditorThemeMode;
}

export function isDocxEditorAttribute(name: string): name is DocxEditorAttribute {
	return (DOCX_EDITOR_ATTRIBUTES as readonly string[]).includes(name);
}

/** attributeChangedCallback body: attribute -> property. */
export function applyAttribute(
	target: AttributeProperties,
	name: DocxEditorAttribute,
	value: string | null,
): void {
	if (name === 'locale') {
		if (normalizeEditorLocale(value ?? 'en') !== target.locale) target.locale = value ?? 'en';
	} else if (name === 'read-only') {
		if ((value !== null) !== target.readOnly) target.readOnly = value !== null;
	} else if (name === 'file-name') {
		const next = value ?? DEFAULT_FILE_NAME;
		if (next !== target.fileName) target.fileName = next;
	} else if (name === 'theme') {
		const next = normalizeThemeMode(value);
		if (next !== target.theme) target.theme = next;
	} else {
		const next = value || DEFAULT_REVIEW_AUTHOR;
		if (next !== target.reviewAuthor) target.reviewAuthor = next;
	}
}

/** Property setter tail: property -> attribute. No-op where the element has no DOM (SSR). */
export function reflectAttribute(
	element: Element,
	name: DocxEditorAttribute,
	value: string | boolean,
): void {
	if (typeof element.setAttribute !== 'function') return;
	if (typeof value === 'boolean') {
		if (element.hasAttribute(name) !== value) element.toggleAttribute(name, value);
		return;
	}
	const current = element.getAttribute(name);
	if (current === value) return;
	// A differently spelled tag such as "FR" already means this locale; leave the author's markup.
	if (name === 'locale' && current !== null && normalizeEditorLocale(current) === value) return;
	element.setAttribute(name, value);
}
