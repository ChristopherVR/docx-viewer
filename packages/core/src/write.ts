// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph, TextRun } from './model.js';
import { children, first, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { writeParagraphProperties } from './write-paragraph-properties.js';

const twips = (pixels: number): string => String(Math.round(pixels * 15));
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
		(run.bold || run.italic || run.underline || run.fontSize || run.fontFamily || run.color)
	) {
		props = makeW(doc, 'rPr');
		runNode.insertBefore(props, runNode.firstChild);
	}
	if (!props) return;
	if (changed('bold')) setToggle(doc, props, 'b', run.bold === true);
	if (changed('italic')) setToggle(doc, props, 'i', run.italic === true);
	if (changed('underline')) {
		removeChildren(props, 'u');
		if (run.underline) {
			const underline = makeW(doc, 'u');
			setAttribute(underline, 'val', 'single');
			props.appendChild(underline);
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

function createRun(doc: XmlDocument, run: TextRun, base?: TextRun, old?: XmlElement): XmlElement {
	const node = old ?? makeW(doc, 'r');
	setRunProperties(doc, node, run, base);
	for (const child of Array.from(node.childNodes))
		if (child.nodeType !== 1 || (child as XmlElement).localName !== 'rPr') node.removeChild(child);
	const pieces = run.text.split(/(\n|\t)/);
	for (const piece of pieces) {
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

function writeTable(
	doc: XmlDocument,
	table: Extract<Block, { type: 'table' }>,
	node: XmlElement,
	base?: Extract<Block, { type: 'table' }>,
): XmlElement {
	if (base && JSON.stringify(table) === JSON.stringify(base)) return node;
	const oldRows = children(node, 'tr');
	const nextRows: XmlElement[] = [];
	table.rows.forEach((row, ri) => {
		const tr = oldRows[ri] ?? makeW(doc, 'tr');
		const oldCells = children(tr, 'tc');
		const nextCells: XmlElement[] = [];
		row.forEach((cell, ci) => {
			const tc = oldCells[ci] ?? makeW(doc, 'tc');
			const oldParagraphs = children(tc, 'p');
			const nextParagraphs = cell.paragraphs.map((paragraph, pi) =>
				writeParagraph(
					doc,
					paragraph,
					oldParagraphs[pi] ?? makeW(doc, 'p'),
					base?.rows[ri]?.[ci]?.paragraphs.find((candidate) => candidate.id === paragraph.id),
				),
			);
			if (!nextParagraphs.length && !oldParagraphs.length) nextParagraphs.push(makeW(doc, 'p'));
			replaceSlots(doc, tc, oldParagraphs, nextParagraphs);
			nextCells.push(tc);
		});
		replaceSlots(doc, tr, oldCells, nextCells);
		nextRows.push(tr);
	});
	replaceSlots(doc, node, oldRows, nextRows);
	return node;
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
					? writeTable(doc, block, old, base?.type === 'table' ? base : undefined)
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
