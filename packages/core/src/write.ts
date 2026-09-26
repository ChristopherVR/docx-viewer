// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph } from './model.js';
import {
	children,
	elements,
	first,
	getW,
	makeW,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { writeParagraphProperties } from './write-paragraph-properties.js';
import { writeTable as writeTableContent } from './write-table.js';
import { createRun } from './write-run.js';
import { isWordHighlightToken } from './highlight.js';

const twips = (pixels: number): string => String(Math.round(pixels * 15));
function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}

function hasUnsafeInline(paragraph: XmlElement): boolean {
	for (const child of Array.from(paragraph.childNodes)) {
		if (child.nodeType !== 1) continue;
		const element = child as XmlElement;
		if (element.localName === 'pPr') continue;
		if (element.localName !== 'r') return true;
		for (const runChild of Array.from(element.childNodes)) {
			if (runChild.nodeType !== 1) continue;
			const runElement = runChild as XmlElement;
			if (runElement.localName === 'rPr') continue;
			if (!['t', 'tab', 'br', 'cr', 'noBreakHyphen'].includes(runElement.localName)) return true;
		}
	}
	return false;
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
				: ['b', 'i', 'strike', 'u', 'highlight', 'vertAlign', 'sz', 'color'].includes(
							property.localName,
					  )
					? ['val']
					: [];
		if (hasUnexpectedAttributes(property, allowed) || elements(property).length > 0) return true;
		const value = getW(property, 'val');
		if (property.localName === 'highlight' && value && !isWordHighlightToken(value)) return true;
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

function rejectUnsafeRunSegmentation(
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

function writeParagraph(
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base?: Paragraph,
): XmlElement {
	if (base && JSON.stringify(paragraph) === JSON.stringify(base)) return node;
	if (hasUnsafeInline(node))
		throw new Error(
			`Cannot edit paragraph ${paragraph.id}: it contains inline OOXML that this editor cannot safely relocate. The original DOCX package remains unchanged.`,
		);
	let pPr = first(node, 'pPr');
	if (!pPr) {
		pPr = makeW(doc, 'pPr');
		node.insertBefore(pPr, node.firstChild);
	}
	if (!base || paragraph.align !== base.align) {
		removeChildren(pPr, 'jc');
		if (paragraph.align) {
			const align = makeW(doc, 'jc');
			setAttribute(align, 'val', paragraph.align === 'justify' ? 'both' : paragraph.align);
			pPr.appendChild(align);
		}
	}
	if (!base || paragraph.style !== base.style) {
		removeChildren(pPr, 'pStyle');
		if (paragraph.style) {
			const style = makeW(doc, 'pStyle');
			setAttribute(style, 'val', paragraph.style);
			pPr.appendChild(style);
		}
	}
	writeParagraphProperties(doc, pPr, paragraph, base);
	const oldRuns = children(node, 'r');
	rejectUnsafeRunSegmentation(paragraph, base, oldRuns);
	for (const run of oldRuns) node.removeChild(run);
	const newRuns = paragraph.runs.map((run, i) => createRun(doc, run, base?.runs[i], oldRuns[i]));
	let anchor: any = pPr;
	for (const run of newRuns) {
		node.insertBefore(run, anchor.nextSibling);
		anchor = run;
	}
	if (!pPr.childNodes.length) node.removeChild(pPr);
	return node;
}

function createParagraph(doc: XmlDocument, paragraph: Paragraph): XmlElement {
	const node = makeW(doc, 'p');
	return writeParagraph(doc, paragraph, node);
}

function createTable(doc: XmlDocument, table: Extract<Block, { type: 'table' }>): XmlElement {
	const node = makeW(doc, 'tbl');
	for (const row of table.rows) {
		const tr = makeW(doc, 'tr');
		for (const cell of row) {
			const tc = makeW(doc, 'tc');
			for (const paragraph of cell.paragraphs) tc.appendChild(createParagraph(doc, paragraph));
			if (!cell.paragraphs.length) tc.appendChild(makeW(doc, 'p'));
			tr.appendChild(tc);
		}
		node.appendChild(tr);
	}
	return node;
}

function replaceSlots(
	doc: XmlDocument,
	parent: XmlElement,
	oldNodes: XmlElement[],
	nextNodes: XmlElement[],
	insertionAnchor: any = null,
): void {
	const placeholders = oldNodes.map((old) => {
		const marker = doc.createComment('docx-block-slot');
		parent.replaceChild(marker, old);
		return marker;
	});
	const shared = Math.min(placeholders.length, nextNodes.length);
	for (let i = 0; i < shared; i++) parent.replaceChild(nextNodes[i], placeholders[i]);
	for (const marker of placeholders.slice(shared)) parent.removeChild(marker);
	for (const node of nextNodes.slice(shared)) parent.insertBefore(node, insertionAnchor);
}

function baseMap(blocks: Block[]): Map<string, Block> {
	return new Map(blocks.map((block) => [block.id, block]));
}
function originalNodes(body: XmlElement): XmlElement[] {
	return Array.from(body.childNodes).filter(
		(node: any) =>
			node.nodeType === 1 &&
			(node as XmlElement).namespaceURI === WORD_NS &&
			['p', 'tbl'].includes((node as XmlElement).localName),
	) as XmlElement[];
}

export function applyModel(doc: XmlDocument, model: DocumentModel, original: Block[]): void {
	const body = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const oldNodes = originalNodes(body);
	const oldById = new Map<string, XmlElement>();
	original.forEach((block, index) => {
		if (oldNodes[index]) oldById.set(block.id, oldNodes[index]);
	});
	const bases = baseMap(original);
	const output: XmlElement[] = [];
	for (const block of model.blocks) {
		const old = oldById.get(block.id);
		const base = bases.get(block.id);
		if (block.type === 'paragraph')
			output.push(
				old
					? writeParagraph(doc, block, old, base?.type === 'paragraph' ? base : undefined)
					: createParagraph(doc, block),
			);
		else
			output.push(
				old
					? writeTableContent(
							doc,
							block,
							old,
							base?.type === 'table' ? base : undefined,
							writeParagraph,
							replaceSlots,
						)
					: createTable(doc, block),
			);
	}
	const slots = originalNodes(body);
	const sectionAnchor = children(body, 'sectPr').at(-1) ?? null;
	replaceSlots(doc, body, slots, output, sectionAnchor);
	let section = children(body, 'sectPr').at(-1);
	if (!section) {
		section = makeW(doc, 'sectPr');
		body.appendChild(section);
	}
	let size = first(section, 'pgSz');
	if (!size) {
		size = makeW(doc, 'pgSz');
		section.insertBefore(size, section.firstChild);
	}
	setAttribute(size, 'w', twips(model.page.width));
	setAttribute(size, 'h', twips(model.page.height));
	let margins = first(section, 'pgMar');
	if (!margins) {
		margins = makeW(doc, 'pgMar');
		section.appendChild(margins);
	}
	for (const side of ['Top', 'Right', 'Bottom', 'Left'] as const)
		setAttribute(margins, side.toLowerCase(), twips(model.page[`margin${side}`]));
}
