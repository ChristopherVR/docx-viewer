// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph, SectionProperties } from './model.js';
import { children, first, getW, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { writeParagraphProperties } from './write-paragraph-properties.js';
import { orderParagraphProperties, writeTabStops } from './tab-stops.js';
import { commentContinuations, continuationKey } from './comment-spans.js';
import { annotationIdAllocator, newMoveRangeIds, type CommentSpans } from './write-ranges.js';
import { reconcileBookmarks } from './bookmarks.js';
import { writeNumberingProperties } from './numbering-write.js';
import { writeTable as writeTableContent } from './write-table.js';
import { buildNewTableProperties, tableColumnCount } from './table-defaults.js';
import { buildCellProperties, buildRowProperties } from './table-cell-write.js';
import { orderSectionProperties } from './element-order.js';
import { applySectionEdits } from './write-sections.js';
import { writeParagraphMarkRevision } from './write-revisions.js';
import { paragraphJustification } from './paragraph-alignment.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import {
	buildInlineContent,
	collectInlineSlots,
	replaceableInlineChildren,
} from './write-inline.js';
import type { DocPrIdAllocator } from './docpr-ids.js';
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
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const continuation = spans?.next.get(paragraph.id);
	const ranges = spans && {
		...(continuation ?? { before: new Set<string>(), after: new Set<string>() }),
		rangeIds: spans.rangeIds,
	};
	const sameRanges =
		continuationKey(continuation) === continuationKey(spans?.original.get(paragraph.id));
	if (base && sameRanges && JSON.stringify(paragraph) === JSON.stringify(base)) return node;
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
	if (
		!base ||
		paragraph.align !== base.align ||
		paragraph.justification !== base.justification ||
		paragraph.direction !== base.direction
	) {
		removeChildren(pPr, 'jc');
		const jc = paragraphJustification(paragraph);
		if (jc) {
			const align = makeW(doc, 'jc');
			setAttribute(align, 'val', jc);
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
	writeTabStops(doc, pPr, paragraph, base);
	orderParagraphProperties(pPr);
	rejectUnsafeRunSegmentation(
		paragraph,
		base,
		slots.flatMap((slot) => (slot.element ? [slot.element] : [])),
	);
	// Bookmarks are preserved but not repositioned precisely: an edited paragraph's bookmarks move
	// to its boundaries (starts right after pPr, ends at the close) instead of their exact original
	// run offsets, which this run-level model does not track.
	const { starts: bookmarkStarts, ends: bookmarkEnds } = reconcileBookmarks(
		doc,
		children(node, 'bookmarkStart'),
		children(node, 'bookmarkEnd'),
		paragraph.bookmarks ?? [],
		base?.bookmarks ?? [],
		() => spans?.nextId() ?? '0',
	);
	for (const bookmark of [...children(node, 'bookmarkStart'), ...children(node, 'bookmarkEnd')])
		node.removeChild(bookmark);
	for (const oldNode of replaceableInlineChildren(node)) node.removeChild(oldNode);
	const newNodes = buildInlineContent(doc, paragraph.runs, base?.runs, slots, allocator, ranges);
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
	allocator: RelationshipAllocator | undefined,
	spans?: CommentSpans,
): XmlElement {
	const node = makeW(doc, 'p');
	return writeParagraphImpl(doc, paragraph, node, undefined, allocator, spans);
}

function createTable(
	doc: XmlDocument,
	table: Extract<Block, { type: 'table' }>,
	allocator: RelationshipAllocator | undefined,
	contentWidthTwips: number,
	spans?: CommentSpans,
): XmlElement {
	const node = makeW(doc, 'tbl');
	const { tblPr, tblGrid } = buildNewTableProperties(doc, table, contentWidthTwips);
	node.appendChild(tblPr);
	node.appendChild(tblGrid);
	const columns = tableColumnCount(table);
	table.rows.forEach((row, rowIndex) => {
		const tr = makeW(doc, 'tr');
		const spanned = row.reduce((sum, cell) => sum + (cell.gridSpan ?? 1), 0);
		const trPr = buildRowProperties(doc, table.rowProperties?.[rowIndex], columns - spanned);
		if (trPr) tr.appendChild(trPr);
		for (const cell of row) {
			const tc = makeW(doc, 'tc');
			const tcPr = buildCellProperties(doc, cell);
			if (tcPr) tc.appendChild(tcPr);
			for (const paragraph of cell.paragraphs)
				tc.appendChild(createParagraph(doc, paragraph, allocator, spans));
			if (!cell.paragraphs.length) tc.appendChild(makeW(doc, 'p'));
			tr.appendChild(tc);
		}
		node.appendChild(tr);
	});
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

/**
 * Applies `blocks` to the paragraphs and tables directly inside `container` (the document body, or
 * a header, footer or note root), rewriting only blocks that changed relative to `original`.
 * Without an allocator, new pictures and external links are rejected (their relationships would
 * belong to a different package part).
 */
export function applyBlocks(
	doc: XmlDocument,
	container: XmlElement,
	blocks: Block[],
	original: Block[],
	allocator: RelationshipAllocator | undefined,
	contentWidthTwips: number,
	anchor: XmlElement | null = null,
): void {
	const nextId = annotationIdAllocator(doc, blocks);
	const spans: CommentSpans = {
		next: commentContinuations(blocks),
		original: commentContinuations(original),
		rangeIds: newMoveRangeIds(blocks, nextId),
		nextId,
	};
	const boundWriteParagraph = (
		writeDoc: XmlDocument,
		paragraph: Paragraph,
		node: XmlElement,
		base?: Paragraph,
	) => writeParagraphImpl(writeDoc, paragraph, node, base, allocator, spans);
	const oldNodes = originalNodes(container);
	const oldById = new Map<string, XmlElement>();
	original.forEach((block, index) => {
		if (oldNodes[index]) oldById.set(block.id, oldNodes[index]);
	});
	const bases = baseMap(original);
	const output: XmlElement[] = [];
	for (const block of blocks) {
		const old = oldById.get(block.id);
		const base = bases.get(block.id);
		if (block.type === 'paragraph')
			output.push(
				old
					? boundWriteParagraph(doc, block, old, base?.type === 'paragraph' ? base : undefined)
					: createParagraph(doc, block, allocator, spans),
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
					: createTable(doc, block, allocator, contentWidthTwips, spans),
			);
	}
	replaceSlots(doc, container, originalNodes(container), output, anchor);
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
	/** The loaded document's sections, so unchanged sections stay byte-identical. */
	baseSections: SectionProperties[] = [],
	/** Package-wide `wp:docPr` id allocator shared with header, footer and note parts. */
	docPrIds?: DocPrIdAllocator,
): ApplyModelResult {
	const body = Array.from(doc.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const allocator = new RelationshipAllocator(
		[...scanUsedRelationshipIds(doc), ...reservedRelationshipIds],
		docPrIds,
	);
	allocator.docPrIds.reserveFromDocument(doc);
	const contentWidthTwips = Math.round(
		(model.page.width - model.page.marginLeft - model.page.marginRight) * 15,
	);
	applyBlocks(
		doc,
		body,
		model.blocks,
		original,
		allocator,
		contentWidthTwips,
		children(body, 'sectPr').at(-1) ?? null,
	);
	let section = children(body, 'sectPr').at(-1);
	if (!section) {
		section = makeW(doc, 'sectPr');
		body.appendChild(section);
	}
	let size = first(section, 'pgSz');
	if (!size) {
		size = makeW(doc, 'pgSz');
		section.appendChild(size);
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
	// CT_PageMar requires header, footer and gutter; keep a source's values, else Word's defaults.
	for (const [name, fallback] of [
		['header', '720'],
		['footer', '720'],
		['gutter', '0'],
	] as const)
		if (!margins.getAttributeNS(WORD_NS, name)) setAttribute(margins, name, fallback);
	orderSectionProperties(section);
	if (model.sections) applySectionEdits(doc, body, model.blocks, model.sections, baseSections);
	return { newRelationships: allocator.newRelationships };
}
