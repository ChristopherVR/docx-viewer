// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { elements, first, getW, type XmlElement, WORD_NS } from './xml.js';
import { isWordHighlightToken } from './highlight.js';
import { isValidLanguageTag } from './language.js';
import { isWordUnderlineStyle } from './underline.js';
import { isThemeColorToken } from './theme-color.js';
import { parseHalfPoints, parseHexColor, parseSignedTwips } from './simple-types.js';

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
	'rStyle',
	'caps',
	'smallCaps',
	'dstrike',
	'vanish',
	'spacing',
	'shd',
]);
const simpleToggleProperties = ['b', 'i', 'strike', 'caps', 'smallCaps', 'dstrike', 'vanish'];
const allowedRunPropertyAttributes: Record<string, string[]> = {
	rFonts: ['ascii', 'hAnsi', 'asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'cstheme'],
	lang: ['val', 'eastAsia', 'bidi'],
	highlight: ['val'],
	vertAlign: ['val'],
	sz: ['val'],
	rtl: ['val'],
	rStyle: ['val'],
	spacing: ['val'],
	u: ['val', 'color'],
	color: ['val', 'themeColor', 'themeTint', 'themeShade'],
	shd: ['val', 'fill', 'color', 'themeFill', 'themeFillTint', 'themeFillShade'],
};
function hasUnexpectedAttributes(element: XmlElement, allowed: string[]): boolean {
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.namespaceURI === 'http://www.w3.org/2000/xmlns/') continue;
		if (attribute.namespaceURI !== WORD_NS || !allowed.includes(attribute.localName)) return true;
	}
	return false;
}

/** Whether a run's `rPr` contains any property/attribute/value this writer does not model. */
export function runHasUnknownProperties(run: XmlElement): boolean {
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
		const allowed = simpleToggleProperties.includes(property.localName)
			? ['val']
			: (allowedRunPropertyAttributes[property.localName] ?? []);
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
			!isWordUnderlineStyle(value) &&
			!['0', 'false', 'off', '1', 'true', 'on'].includes(value)
		)
			return true;
		if (property.localName === 'u') {
			const color = getW(property, 'color');
			if (color && parseHexColor(color) === undefined) return true;
		}
		if (property.localName === 'sz' && value && parseHalfPoints(value) === undefined) return true;
		if (property.localName === 'spacing' && value && parseSignedTwips(value) === undefined)
			return true;
		if (property.localName === 'color' || property.localName === 'shd') {
			const fillValue = property.localName === 'shd' ? getW(property, 'fill') : value;
			if (fillValue && parseHexColor(fillValue) === undefined) return true;
			const themeValue = getW(property, property.localName === 'shd' ? 'themeFill' : 'themeColor');
			if (themeValue && !isThemeColorToken(themeValue)) return true;
		}
	}
	return false;
}
