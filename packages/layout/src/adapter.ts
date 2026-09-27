import type { Block, DocumentModel, Paragraph, Table, TextRun } from '@christophervr/docx-core';
import { resolveParagraphFormatting } from '@christophervr/docx-core';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutParagraph,
	LayoutRun,
	LayoutTable,
} from './input.js';

const KEEP_TOGETHER_NOTE =
	'Paragraph keepNext/keepLines/widowControl overrides, pageBreakBefore, and contextualSpacing are not yet represented in the document model; Word’s defaults (widow/orphan control on, no forced keep-together) are used for every paragraph.';
const TABLE_ROW_NOTE =
	'Table row "keep together" (cantSplit) and repeating header rows (tblHeader) are not yet represented in the document model; every row may split across a page.';
const BREAK_NOTE =
	'Explicit page and column breaks are not distinguished from ordinary line breaks in the document model; they are laid out as in-paragraph line breaks rather than starting a new page/column.';

/**
 * Optional per-section geometry a future sections/headers-footers model
 * extension may add to `DocumentModel`. Recognized only when present; a
 * model without it lays out as a single section using `model.page`. This is
 * the seam the sections agent's model is expected to plug into at merge time.
 */
export interface ModelSectionHint {
	startBlockIndex?: number;
	endBlockIndex?: number;
	page?: Partial<{
		width: number;
		height: number;
		marginTop: number;
		marginRight: number;
		marginBottom: number;
		marginLeft: number;
	}>;
	columns?: { count?: number; gapPx?: number };
	break?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
}
export type ModelWithSections = DocumentModel &
	Partial<{ sections: ReadonlyArray<ModelSectionHint> }>;

/** Converts a `docx-core` `DocumentModel` into this engine's own input contract. */
export function adaptDocumentModel(
	model: DocumentModel,
	note: (message: string) => void = () => {},
): LayoutDocumentInput {
	const catalog = model.paragraphStyles;
	const noted = new Set<string>();
	const reportOnce = (message: string) => {
		if (!noted.has(message)) {
			noted.add(message);
			note(message);
		}
	};

	function adaptRun(run: TextRun): LayoutRun {
		if (/[\n\t]/.test(run.text)) reportOnce(BREAK_NOTE);
		return {
			text: run.text,
			bold: run.bold,
			italic: run.italic,
			fontFamily: run.fontFamily,
			fontSizePt: run.fontSize,
		};
	}
	function adaptParagraph(paragraph: Paragraph): LayoutParagraph {
		const resolved = catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph;
		reportOnce(KEEP_TOGETHER_NOTE);
		return {
			kind: 'paragraph',
			id: paragraph.id,
			runs: paragraph.runs.map(adaptRun),
			align: resolved.align,
			direction: resolved.direction,
			spacingBeforeTwips: resolved.spacingBeforeTwips,
			spacingAfterTwips: resolved.spacingAfterTwips,
			lineSpacingTwips: resolved.lineSpacingTwips,
			lineSpacingRule: resolved.lineSpacingRule,
			indentLeftTwips: resolved.indentLeftTwips,
			indentRightTwips: resolved.indentRightTwips,
			indentStartTwips: resolved.indentStartTwips,
			indentEndTwips: resolved.indentEndTwips,
			firstLineTwips: resolved.firstLineTwips,
			hangingTwips: resolved.hangingTwips,
			styleId: paragraph.style,
		};
	}
	function adaptTable(table: Table): LayoutTable {
		reportOnce(TABLE_ROW_NOTE);
		return {
			kind: 'table',
			id: table.id,
			rows: table.rows.map((row) => ({
				cells: row.map((cell) => ({ paragraphs: cell.paragraphs.map(adaptParagraph) })),
			})),
		};
	}
	function adaptBlock(block: Block): LayoutBlock {
		return block.type === 'table' ? adaptTable(block) : adaptParagraph(block);
	}

	const blocks = model.blocks.map(adaptBlock);
	const sectionHints = (model as ModelWithSections).sections;
	if (!sectionHints || !sectionHints.length) {
		return {
			sections: [
				{
					page: {
						widthPx: model.page.width,
						heightPx: model.page.height,
						marginTopPx: model.page.marginTop,
						marginRightPx: model.page.marginRight,
						marginBottomPx: model.page.marginBottom,
						marginLeftPx: model.page.marginLeft,
					},
					blocks,
				},
			],
		};
	}
	return {
		sections: sectionHints.map((hint, index) => ({
			page: {
				widthPx: hint.page?.width ?? model.page.width,
				heightPx: hint.page?.height ?? model.page.height,
				marginTopPx: hint.page?.marginTop ?? model.page.marginTop,
				marginRightPx: hint.page?.marginRight ?? model.page.marginRight,
				marginBottomPx: hint.page?.marginBottom ?? model.page.marginBottom,
				marginLeftPx: hint.page?.marginLeft ?? model.page.marginLeft,
			},
			columns: hint.columns
				? { count: hint.columns.count ?? 1, gapPx: hint.columns.gapPx ?? 0 }
				: undefined,
			break: index === 0 ? undefined : hint.break,
			blocks: blocks.slice(hint.startBlockIndex ?? 0, hint.endBlockIndex ?? blocks.length),
		})),
	};
}
