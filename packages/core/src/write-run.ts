// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { isWordHighlightToken } from './highlight.js';
import { isValidLanguageTag } from './language.js';

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
			run.color)
	) {
		props = makeW(doc, 'rPr');
		runNode.insertBefore(props, runNode.firstChild);
	}
	if (!props) return;
	if (changed('bold')) setToggle(doc, props, 'b', run.bold === true);
	if (changed('italic')) setToggle(doc, props, 'i', run.italic === true);
	if (changed('strike')) {
		setToggle(doc, props, 'strike', run.strike === true);
		removeChildren(props, 'dstrike');
	}
	if (changed('underline')) {
		removeChildren(props, 'u');
		if (run.underline) {
			const underline = makeW(doc, 'u');
			setAttribute(underline, 'val', 'single');
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
	if (changed('fontFamily')) {
		removeChildren(props, 'rFonts');
		if (run.fontFamily) {
			const fonts = makeW(doc, 'rFonts');
			setAttribute(fonts, 'ascii', run.fontFamily);
			setAttribute(fonts, 'hAnsi', run.fontFamily);
			props.appendChild(fonts);
		}
	}
	if (changed('color')) {
		removeChildren(props, 'color');
		if (run.color) {
			const color = makeW(doc, 'color');
			setAttribute(color, 'val', run.color.replace(/^#/, ''));
			props.appendChild(color);
		}
	}
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
