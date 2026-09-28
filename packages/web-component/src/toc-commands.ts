import {
	buildTableOfContents,
	DEFAULT_TOC_INSTRUCTION,
	findTableOfContents,
	tocBookmarks,
	tocEntries,
	tocHyperlinks,
	type Block,
	type DocumentModel,
	type Paragraph,
} from '@christophervr/docx-core';
import {
	createCanvasMeasurer,
	layoutDocumentModel,
	type TextMeasurer,
} from '@christophervr/docx-layout';
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
		for (const column of page.columns)
			for (const block of column.blocks)
				if (!result.has(block.blockId)) result.set(block.blockId, numbers[index]);
	});
	return result;
}

const contentWidthTwips = (model: DocumentModel) =>
	Math.round((model.page.width - model.page.marginLeft - model.page.marginRight) * 15);

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

/** Inserts a TOC before the block holding the caret, replacing it when it is an empty paragraph. */
export function insertTableOfContents(
	view: EditorView,
	model: DocumentModel,
	measurer?: TextMeasurer,
): void {
	const { $from } = view.state.selection;
	const index = $from.index(0);
	const current = view.state.doc.maybeChild(index);
	const replace = current?.type.name === 'paragraph' && current.content.size === 0;
	const end = replace ? index : index - 1;
	replaceBlocks(view, model, index, end, tocParagraphs(model, index, end, undefined, measurer));
}

/** Rebuilds the document's TOC with current headings and page numbers; false when there is none. */
export function updateTableOfContents(
	view: EditorView,
	model: DocumentModel,
	measurer?: TextMeasurer,
): boolean {
	const found = findTableOfContents(model.blocks);
	if (!found) return false;
	const toc = tocParagraphs(model, found.start, found.end, found.instruction, measurer);
	toc.paragraphs[0].runs.unshift(...found.before);
	toc.paragraphs.at(-1)!.runs.push(...found.after);
	replaceBlocks(view, model, found.start, found.end, toc);
	return true;
}
