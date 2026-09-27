import type {
	Block,
	DocumentModel,
	Paragraph,
	SectionProperties,
	Table,
	TextRun,
} from '@christophervr/docx-core';
import { resolveParagraphFormatting } from '@christophervr/docx-core';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutParagraph,
	LayoutRun,
	LayoutSection,
	LayoutTable,
} from './input.js';

const KEEP_TOGETHER_NOTE =
	'Paragraph keepNext/keepLines/widowControl overrides and contextualSpacing are not yet represented in the document model; Word’s defaults (widow/orphan control on, no forced keep-together) are used for every paragraph.';
const TABLE_ROW_NOTE =
	'Table row "keep together" (cantSplit) and repeating header rows (tblHeader) are not yet represented in the document model; every row may split across a page.';
const NEXT_COLUMN_NOTE =
	'A "next column" section break is laid out as a continuous section break.';

const twipsToPx = (twips: number): number => twips / 15;

function sectionBreak(
	type: SectionProperties['type'],
	note: (message: string) => void,
): LayoutSection['break'] {
	if (type === 'nextColumn') {
		note(NEXT_COLUMN_NOTE);
		return 'continuous';
	}
	return type;
}

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
		return {
			text: run.text,
			bold: run.bold,
			italic: run.italic,
			fontFamily: run.fontFamily,
			fontSizePt: run.fontSize,
			...(run.break ? { breakAfter: run.break } : {}),
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
			...(paragraph.pageBreakBefore ? { pageBreakBefore: true } : {}),
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
	const sections = model.sections;
	if (!sections || !sections.length) {
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
	const indexById = new Map(model.blocks.map((block, index) => [block.id, index]));
	let start = 0;
	const result: LayoutSection[] = sections.map((section, index) => {
		const last = index === sections.length - 1;
		const endIndex = last ? blocks.length : (indexById.get(section.endsAtBlockId) ?? start - 1) + 1;
		const end = Math.max(start, endIndex);
		const slice = blocks.slice(start, end);
		start = end;
		return {
			page: {
				widthPx: twipsToPx(section.pageWidthTwips),
				heightPx: twipsToPx(section.pageHeightTwips),
				marginTopPx: twipsToPx(section.marginTopTwips),
				marginRightPx: twipsToPx(section.marginRightTwips),
				marginBottomPx: twipsToPx(section.marginBottomTwips),
				marginLeftPx: twipsToPx(section.marginLeftTwips),
			},
			...(section.columns.count > 1
				? {
						columns: {
							count: section.columns.count,
							gapPx: twipsToPx(section.columns.spacingTwips ?? 720),
						},
					}
				: {}),
			...(index === 0 ? {} : { break: sectionBreak(section.type, reportOnce) }),
			blocks: slice,
		};
	});
	return { sections: result };
}
