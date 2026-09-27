// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph } from './model.js';
import { children, first, getW, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { writeParagraphProperties } from './write-paragraph-properties.js';
import { writeNumberingProperties } from './numbering-write.js';
import { writeTable as writeTableContent } from './write-table.js';
import { writeParagraphMarkRevision } from './write-revisions.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import {
	buildInlineContent,
	collectInlineSlots,
	replaceableInlineChildren,
} from './write-inline.js';
import {
	RelationshipAllocator,
	scanUsedRelationshipIds,
	type NewRelationship,
} from './relationship-allocator.js';

const twips = (pixels: number): string => String(Math.round(pixels * 15));
function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
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

function writeParagraphImpl(
	doc: XmlDocument,
	paragraph: Paragraph,
	node: XmlElement,
	base: Paragraph | undefined,
	allocator: RelationshipAllocator,
): XmlElement {
	if (base && JSON.stringify(paragraph) === JSON.stringify(base)) return node;
	const slots = collectInlineSlots(node);
	if (!slots)
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
	writeNumberingProperties(doc, pPr, paragraph, base);
	writeParagraphMarkRevision(doc, pPr, paragraph, base);
	removeChildren(pPr, 'pPrChange');
	rejectUnsafeRunSegmentation(
		paragraph,
		base,
		slots.map((slot) => slot.element),
	);
	// Bookmarks are preserved but not repositioned precisely: an edited paragraph's bookmarks move
	// to its boundaries (starts right after pPr, ends at the close) instead of their exact original
	// run offsets, which this run-level model does not track.
	const bookmarkStarts = children(node, 'bookmarkStart');
	const bookmarkEnds = children(node, 'bookmarkEnd');
	for (const bookmark of [...bookmarkStarts, ...bookmarkEnds]) node.removeChild(bookmark);
	for (const oldNode of replaceableInlineChildren(node)) node.removeChild(oldNode);
	const newNodes = buildInlineContent(doc, paragraph.runs, base?.runs, slots, allocator);
	let anchor: XmlElement = pPr;
	for (const bookmark of bookmarkStarts) {
		node.insertBefore(bookmark, anchor.nextSibling);
		anchor = bookmark;
	}
	for (const run of newNodes) {
		node.insertBefore(run, anchor.nextSibling);
		anchor = run;
	}
	for (const bookmark of bookmarkEnds) {
		node.insertBefore(bookmark, anchor.nextSibling);
		anchor = bookmark;
	}
	if (!pPr.childNodes.length) node.removeChild(pPr);
	return node;
}

function createParagraph(
	doc: XmlDocument,
	paragraph: Paragraph,
	allocator: RelationshipAllocator,
): XmlElement {
	const node = makeW(doc, 'p');
	return writeParagraphImpl(doc, paragraph, node, undefined, allocator);
}

function createTable(
	doc: XmlDocument,
	table: Extract<Block, { type: 'table' }>,
	allocator: RelationshipAllocator,
): XmlElement {
	const node = makeW(doc, 'tbl');
	for (const row of table.rows) {
		const tr = makeW(doc, 'tr');
		for (const cell of row) {
			const tc = makeW(doc, 'tc');
			for (const paragraph of cell.paragraphs)
				tc.appendChild(createParagraph(doc, paragraph, allocator));
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

export interface ApplyModelResult {
	/** New hyperlink/image relationships save.ts must add to word/_rels/document.xml.rels. */
	newRelationships: readonly NewRelationship[];
}

export function applyModel(
	doc: XmlDocument,
	model: DocumentModel,
	original: Block[],
	/** Relationship ids already declared in word/_rels/document.xml.rels (styles, numbering, ...). */
	reservedRelationshipIds: Iterable<string> = [],
): ApplyModelResult {
	const body = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const allocator = new RelationshipAllocator([
		...scanUsedRelationshipIds(doc),
		...reservedRelationshipIds,
	]);
	const boundWriteParagraph = (
		writeDoc: XmlDocument,
		paragraph: Paragraph,
		node: XmlElement,
		base?: Paragraph,
	) => writeParagraphImpl(writeDoc, paragraph, node, base, allocator);
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
					? boundWriteParagraph(doc, block, old, base?.type === 'paragraph' ? base : undefined)
					: createParagraph(doc, block, allocator),
			);
		else
			output.push(
				old
					? writeTableContent(
							doc,
							block,
							old,
							base?.type === 'table' ? base : undefined,
							boundWriteParagraph,
							replaceSlots,
						)
					: createTable(doc, block, allocator),
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
	return { newRelationships: allocator.newRelationships };
}
