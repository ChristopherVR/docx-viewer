import {
	buildTableOfContents,
	DEFAULT_TOC_INSTRUCTION,
	findTableOfContents,
	tocBookmarks,
	tableOfFiguresInstruction,
	tocCaptionLabel,
	tocEntries,
	tocHyperlinks,
	twipsFromPixels,
	type Block,
	type DocumentModel,
	type Paragraph,
} from 'docx-core';
import { layoutDocumentModel, type TextMeasurer } from 'ooxml-core/docx/layout';
import { createCanvasMeasurer } from './canvas-measurer';
import { closeHistory } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { modelToDoc } from './model-adapter';
import { pageNumbers } from './print-header-footer';

/** Word page number text of the page where each top-level block starts, from Print Layout pagination. */
export function blockPageNumbers(
	model: DocumentModel,
	measurer: TextMeasurer = createCanvasMeasurer(),
): Map<string, string> {
	const { pages } = layoutDocumentModel(model, measurer);
	const numbers = pageNumbers(model, pages);
	const result = new Map<string, string>();
	pages.forEach((page, index) => {
		const number = numbers[index];
		if (number === undefined) return;
		for (const column of page.columns)
			for (const block of column.blocks)
				if (!result.has(block.blockId)) result.set(block.blockId, number);
	});
	return result;
}

const contentWidthTwips = (model: DocumentModel) =>
	twipsFromPixels(model.page.width - model.page.marginLeft - model.page.marginRight);

/**
 * Builds TOC paragraphs to place at `start..end` (inclusive, or `end = start - 1` to insert),
 * numbering pages from a layout of the document with the TOC in place, as Word does.
 */
function tocParagraphs(
	model: DocumentModel,
	start: number,
	end: number,
	instruction: string | undefined,
	measurer?: TextMeasurer,
): { paragraphs: Paragraph[]; added: Map<string, string> } {
	let counter = 0;
	const code = instruction ?? DEFAULT_TOC_INSTRUCTION;
	const linked = tocHyperlinks(code) ? tocBookmarks(model, tocEntries(model, code)) : undefined;
	const options = {
		instruction: code,
		contentWidthTwips: contentWidthTwips(model),
		newId: () => `dve-toc-draft-${++counter}`,
		...(linked ? { bookmarks: linked.bookmarks } : {}),
	};
	const draft = buildTableOfContents(model, options);
	const blocks: Block[] = [...model.blocks];
	blocks.splice(start, end - start + 1, ...draft);
	const numbers = blockPageNumbers({ ...model, blocks }, measurer);
	// Empty ids are assigned by the editor's id repair (or collaboration ids) when inserted.
	return {
		paragraphs: buildTableOfContents(model, { ...options, pageNumbers: numbers, newId: () => '' }),
		added: linked?.added ?? new Map(),
	};
}

function replaceBlocks(
	view: EditorView,
	model: DocumentModel,
	start: number,
	end: number,
	{ paragraphs, added }: { paragraphs: Paragraph[]; added: Map<string, string> },
) {
	const { doc } = view.state;
	let from = 0;
	for (let index = 0; index < start; index++) from += doc.child(index).nodeSize;
	let to = from;
	for (let index = start; index <= end; index++) to += doc.child(index).nodeSize;
	const content = modelToDoc({ ...model, blocks: paragraphs }).content;
	const tr = view.state.tr.replaceWith(from, to, content);
	// Headings get the `_Toc` bookmarks their entries link to, in the same undoable step.
	tr.doc.descendants((node, pos) => {
		const name = node.type.name === 'paragraph' ? added.get(String(node.attrs.id)) : undefined;
		if (name)
			tr.setNodeMarkup(pos, undefined, {
				...node.attrs,
				bookmarks: [...((node.attrs.bookmarks as string[]) ?? []), name],
			});
		return node.type.name !== 'paragraph';
	});
	view.dispatch(closeHistory(tr).scrollIntoView());
}

/** Selects the table of figures of `label` (case-insensitively), or heading TOCs when absent. */
const figuresFilter = (label: string | undefined) =>
	label
		? (instruction: string) => tocCaptionLabel(instruction)?.toLowerCase() === label.toLowerCase()
		: undefined;

/** Inserts a TOC before the block holding the caret, replacing it when it is an empty paragraph. */
export function insertTableOfContents(
	view: EditorView,
	model: DocumentModel,
	measurer?: TextMeasurer,
	levels = 3,
	captionLabel?: string,
): void {
	const { $from } = view.state.selection;
	const index = $from.index(0);
	const current = view.state.doc.maybeChild(index);
	const replace = current?.type.name === 'paragraph' && current.content.size === 0;
	const end = replace ? index : index - 1;
	const depth = Math.min(9, Math.max(1, Math.floor(levels)));
	const instruction = captionLabel
		? tableOfFiguresInstruction(captionLabel)
		: ` TOC \\o "1-${depth}" \\h \\z \\u `;
	replaceBlocks(view, model, index, end, tocParagraphs(model, index, end, instruction, measurer));
}

/**
 * Rebuilds the document's TOC with current headings and page numbers (or, with `captionLabel`,
 * its table of figures from the current captions); false when there is none.
 */
export function updateTableOfContents(
	view: EditorView,
	model: DocumentModel,
	measurer?: TextMeasurer,
	captionLabel?: string,
): boolean {
	const found = findTableOfContents(model.blocks, figuresFilter(captionLabel));
	if (!found) return false;
	const toc = tocParagraphs(model, found.start, found.end, found.instruction, measurer);
	const firstParagraph = toc.paragraphs[0];
	const lastParagraph = toc.paragraphs.at(-1);
	if (!firstParagraph || !lastParagraph) return false;
	firstParagraph.runs.unshift(...found.before);
	lastParagraph.runs.push(...found.after);
	replaceBlocks(view, model, found.start, found.end, toc);
	return true;
}

/**
 * Removes the document's table of contents (Word's Remove Table of Contents). Text that shared a
 * paragraph with the field is kept. Returns false when there is no TOC to remove.
 */
export function removeTableOfContents(
	view: EditorView,
	model: DocumentModel,
	captionLabel?: string,
): boolean {
	const found = findTableOfContents(model.blocks, figuresFilter(captionLabel));
	if (!found || !view.editable) return false;
	const { doc } = view.state;
	let from = 0;
	for (let index = 0; index < found.start; index++) from += doc.child(index).nodeSize;
	let to = from;
	for (let index = found.start; index <= found.end; index++) to += doc.child(index).nodeSize;
	const leftover = [...found.before, ...found.after].filter((run) => run.text);
	const content = leftover.length
		? modelToDoc({
				...model,
				blocks: [{ type: 'paragraph', id: '', runs: leftover } as Paragraph],
			}).content
		: undefined;
	const tr = content
		? view.state.tr.replaceWith(from, to, content)
		: view.state.tr.delete(from, to);
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
