// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Guards that decide whether a paragraph's run XML can be safely regenerated without dropping
// unmodeled run properties (rPr we don't understand) or unmodeled inline content.
import type { Paragraph } from './model.js';
import { elements, first, getW, type XmlElement, WORD_NS } from './xml.js';
import { isWordHighlightToken } from './highlight.js';
import { hasSpecialBreak } from './breaks.js';
import { isValidLanguageTag } from './language.js';

/** Validates a single run's children; `w:hyperlink` wrapping and bookmarks are handled by the caller. */
export function hasUnsafeRunContent(run: XmlElement): boolean {
	const kinds = new Set<string>();
	for (const child of Array.from(run.childNodes)) {
		if (child.nodeType !== 1) continue;
		const element = child as XmlElement;
		if (element.localName === 'rPr') continue;
		if (element.localName === 'br' && hasSpecialBreak(element)) return true;
		if (element.localName === 'cr' && element.attributes.length > 0) return true;
		if (!['t', 'tab', 'br', 'cr', 'noBreakHyphen', 'drawing', 'pict'].includes(element.localName))
			return true;
		kinds.add(element.localName);
	}
	// A run modeled as text or as a picture, never both; mixed content isn't represented.
	return (kinds.has('drawing') || kinds.has('pict')) && kinds.size > 1;
}

const modeledRunProperties = new Set([
	'b',
	'i',
	'strike',
	'u',
	'highlight',
	'vertAlign',
	'sz',
	'rFonts',
	'color',
	'lang',
	'rtl',
]);
function hasUnexpectedAttributes(element: XmlElement, allowed: string[]): boolean {
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS || !allowed.includes(attribute.localName)) return true;
	}
	return false;
}
function runHasUnknownProperties(run: XmlElement): boolean {
	const properties = first(run, 'rPr');
	if (!properties) return false;
	if (hasUnexpectedAttributes(properties, [])) return true;
	for (const node of Array.from(properties.childNodes)) {
		if (node.nodeType !== 1) {
			if (node.nodeType === 3 && node.textContent?.trim()) return true;
			continue;
		}
		const property = node as XmlElement;
		if (property.namespaceURI !== WORD_NS || !modeledRunProperties.has(property.localName))
			return true;
		const allowed =
			property.localName === 'rFonts'
				? ['ascii', 'hAnsi']
				: property.localName === 'lang'
					? ['val', 'eastAsia', 'bidi']
					: ['b', 'i', 'strike', 'u', 'highlight', 'vertAlign', 'sz', 'color', 'rtl'].includes(
								property.localName,
						  )
						? ['val']
						: [];
		if (hasUnexpectedAttributes(property, allowed) || elements(property).length > 0) return true;
		const value = getW(property, 'val');
		if (property.localName === 'highlight' && value && !isWordHighlightToken(value)) return true;
		if (
			property.localName === 'rtl' &&
			value &&
			!['1', 'true', 'on', '0', 'false', 'off', 'no'].includes(value)
		)
			return true;
		if (
			property.localName === 'lang' &&
			['val', 'eastAsia', 'bidi'].some((key) => {
				const language = getW(property, key);
				return language !== undefined && !isValidLanguageTag(language);
			})
		)
			return true;
		if (property.localName === 'vertAlign' && value !== 'superscript' && value !== 'subscript')
			return true;
		if (
			property.localName === 'u' &&
			value &&
			!['single', 'none', '0', 'false', 'off', '1', 'true', 'on'].includes(value)
		)
			return true;
		if (property.localName === 'sz' && value && !/^\d+$/.test(value)) return true;
		if (property.localName === 'color' && value && !/^[0-9a-f]{6}$/i.test(value)) return true;
	}
	return false;
}

export function rejectUnsafeRunSegmentation(
	paragraph: Paragraph,
	base: Paragraph | undefined,
	oldRuns: XmlElement[],
): void {
	if (!base || !oldRuns.some(runHasUnknownProperties)) return;
	const sameText =
		paragraph.runs.map((run) => run.text).join('') === base.runs.map((run) => run.text).join('');
	const sameBoundaries =
		paragraph.runs.length === base.runs.length &&
		paragraph.runs.every((run, index) => run.text === base.runs[index]?.text);
	if (sameBoundaries || (!sameText && paragraph.runs.length === base.runs.length)) return;
	throw new Error(
		`Cannot edit paragraph ${paragraph.id}: changing run boundaries could drop unsupported run properties. The original DOCX package remains unchanged.`,
	);
}
