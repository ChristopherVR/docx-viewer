// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { isWordHighlightToken } from './highlight.js';

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
