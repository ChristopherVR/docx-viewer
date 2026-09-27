// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { isWordHighlightToken } from './highlight.js';
import { isValidLanguageTag } from './language.js';
import { setExtendedRunProperties } from './write-run-extra.js';
import { fractionToThemeByte } from './theme-color.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}
function setToggle(doc: XmlDocument, props: XmlElement, local: string, enabled: boolean): void {
	removeChildren(props, local);
	if (enabled) props.appendChild(makeW(doc, local));
}

function setBooleanAttribute(
	doc: XmlDocument,
	props: XmlElement,
	local: string,
	value: boolean | undefined,
): void {
	let element = first(props, local);
	if (value === undefined) {
		removeChildren(props, local);
		return;
	}
	if (!element) {
		element = makeW(doc, local);
		props.appendChild(element);
	}
	setAttribute(element, 'val', value ? '1' : '0');
}

function setLanguageAttribute(
	doc: XmlDocument,
	props: XmlElement,
	attribute: 'val' | 'eastAsia' | 'bidi',
	value: string | undefined,
): void {
	let language = first(props, 'lang');
	if (value !== undefined && value !== '' && !isValidLanguageTag(value))
		throw new Error(`Invalid BCP 47 language tag: ${value}`);
	if (value === undefined || value === '') {
		if (!language) return;
		language.removeAttributeNS(WORD_NS, attribute);
	} else {
		if (!language) {
			language = makeW(doc, 'lang');
			props.appendChild(language);
		}
		setAttribute(language, attribute, value);
	}
	if (language && !language.attributes.length && !language.childNodes.length)
		props.removeChild(language);
}

function setRunProperties(
	doc: XmlDocument,
	runNode: XmlElement,
	run: TextRun,
	base?: TextRun,
): void {
	let props = first(runNode, 'rPr');
	const changed = (key: keyof TextRun): boolean => !base || run[key] !== base[key];
	if (
		!props &&
		(run.bold ||
			run.italic ||
			run.underline ||
			run.strike ||
			run.highlight ||
			run.verticalAlign ||
			run.language !== undefined ||
			run.eastAsiaLanguage !== undefined ||
			run.bidiLanguage !== undefined ||
			run.rtl !== undefined ||
			run.fontSize ||
			run.fontFamily ||
			run.color ||
			run.colorTheme ||
			run.style ||
			run.caps ||
			run.smallCaps ||
			run.doubleStrike ||
			run.vanish ||
			run.underlineStyle ||
			run.characterSpacingTwips !== undefined ||
			run.shadingFill ||
			run.shadingThemeFill ||
			run.fontTheme)
	) {
		props = makeW(doc, 'rPr');
		runNode.insertBefore(props, runNode.firstChild);
	}
	if (!props) return;
	// A rewritten run drops any recorded formatting-change snapshot; reconstructing historical
	// rPrChange diffs is not supported (see model.ts Revision / parse-revisions.ts).
	removeChildren(props, 'rPrChange');
	if (changed('bold')) setToggle(doc, props, 'b', run.bold === true);
	if (changed('italic')) setToggle(doc, props, 'i', run.italic === true);
	if (changed('strike')) {
		setToggle(doc, props, 'strike', run.strike === true);
		removeChildren(props, 'dstrike');
	}
	if (changed('underline') || changed('underlineStyle') || changed('underlineColor')) {
		removeChildren(props, 'u');
		if (run.underline) {
			const underline = makeW(doc, 'u');
			setAttribute(underline, 'val', run.underlineStyle ?? 'single');
			if (run.underlineColor) setAttribute(underline, 'color', run.underlineColor.replace(/^#/, ''));
			props.appendChild(underline);
		}
	}
	if (changed('highlight')) {
		if (run.highlight && !isWordHighlightToken(run.highlight))
			throw new Error(`Unsupported Word highlight token: ${run.highlight}`);
		removeChildren(props, 'highlight');
		if (run.highlight) {
			const highlight = makeW(doc, 'highlight');
			setAttribute(highlight, 'val', run.highlight);
			props.appendChild(highlight);
		}
	}
	if (changed('verticalAlign')) {
		removeChildren(props, 'vertAlign');
		if (run.verticalAlign) {
			const verticalAlign = makeW(doc, 'vertAlign');
			setAttribute(verticalAlign, 'val', run.verticalAlign);
			props.appendChild(verticalAlign);
		}
	}
	if (changed('language')) setLanguageAttribute(doc, props, 'val', run.language);
	if (changed('eastAsiaLanguage'))
		setLanguageAttribute(doc, props, 'eastAsia', run.eastAsiaLanguage);
	if (changed('bidiLanguage')) setLanguageAttribute(doc, props, 'bidi', run.bidiLanguage);
	if (changed('rtl')) setBooleanAttribute(doc, props, 'rtl', run.rtl);
	if (changed('fontSize')) {
		removeChildren(props, 'sz');
		if (run.fontSize !== undefined) {
			const size = makeW(doc, 'sz');
			setAttribute(size, 'val', String(Math.round(run.fontSize * 2)));
			props.appendChild(size);
		}
	}
	if (changed('fontFamily') || changed('fontTheme')) {
		removeChildren(props, 'rFonts');
		if (run.fontFamily || run.fontTheme) {
			const fonts = makeW(doc, 'rFonts');
			if (run.fontFamily) {
				setAttribute(fonts, 'ascii', run.fontFamily);
				setAttribute(fonts, 'hAnsi', run.fontFamily);
			}
			const themeAttribute: Record<'ascii' | 'hAnsi' | 'eastAsia' | 'cs', [string, string]> = {
				ascii: ['asciiTheme', 'Ascii'],
				hAnsi: ['hAnsiTheme', 'HAnsi'],
				eastAsia: ['eastAsiaTheme', 'EastAsia'],
				cs: ['cstheme', 'Bidi'],
			};
			for (const script of ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const) {
				const role = run.fontTheme?.[script];
				const [attribute, suffix] = themeAttribute[script];
				if (role) setAttribute(fonts, attribute, `${role}${suffix}`);
			}
			props.appendChild(fonts);
		}
	}
	if (changed('color') || changed('colorTheme')) {
		removeChildren(props, 'color');
		if (run.color || run.colorTheme) {
			const color = makeW(doc, 'color');
			setAttribute(color, 'val', run.color ? run.color.replace(/^#/, '') : 'auto');
			if (run.colorTheme) {
				setAttribute(color, 'themeColor', run.colorTheme.token);
				if (run.colorTheme.tint !== undefined)
					setAttribute(color, 'themeTint', fractionToThemeByte(run.colorTheme.tint));
				if (run.colorTheme.shade !== undefined)
					setAttribute(color, 'themeShade', fractionToThemeByte(run.colorTheme.shade));
			}
			props.appendChild(color);
		}
	}
	setExtendedRunProperties(doc, props, run, base);
	if (!props.childNodes.length) runNode.removeChild(props);
}

export function createRun(
	doc: XmlDocument,
	run: TextRun,
	base?: TextRun,
	old?: XmlElement,
): XmlElement {
	const node = old ?? makeW(doc, 'r');
	setRunProperties(doc, node, run, base);
	for (const child of Array.from(node.childNodes))
		if (child.nodeType !== 1 || (child as XmlElement).localName !== 'rPr') node.removeChild(child);
	if (run.break) {
		const br = makeW(doc, 'br');
		setAttribute(br, 'type', run.break);
		node.appendChild(br);
		return node;
	}
	if (run.noteReference) {
		const reference = makeW(
			doc,
			run.noteReference.kind === 'footnote' ? 'footnoteReference' : 'endnoteReference',
		);
		setAttribute(reference, 'id', run.noteReference.id);
		node.appendChild(reference);
		return node;
	}
	for (const piece of run.text.split(/(\n|\t)/)) {
		if (piece === '\n') node.appendChild(makeW(doc, 'br'));
		else if (piece === '\t') node.appendChild(makeW(doc, 'tab'));
		else if (piece) {
			const text = makeW(doc, 't');
			if (/^\s|\s$/.test(piece))
				text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
			text.appendChild(doc.createTextNode(piece));
			node.appendChild(text);
		}
	}
	if (!run.text) node.appendChild(makeW(doc, 't'));
	return node;
}
