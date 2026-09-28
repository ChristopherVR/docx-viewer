import type {
	Block,
	DocumentModel,
	Paragraph,
	SectionProperties,
	Table,
	TextRun,
} from '@christophervr/docx-core';
import {
	computeListLabels,
	dateFieldResult,
	displayListLabel,
	type ParagraphListLabel,
	fieldName,
	resolveParagraphFormatting,
	resolveRunFormatting,
	resolveThemeColorReference,
} from '@christophervr/docx-core';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutFloat,
	LayoutParagraph,
	LayoutRun,
	LayoutSection,
	LayoutTable,
} from './input.js';

const KEEP_TOGETHER_NOTE =
	'Paragraph keepNext/keepLines/widowControl overrides and contextualSpacing are not yet represented in the document model; Word’s defaults (widow/orphan control on, no forced keep-together) are used for every paragraph.';
const TABLE_ROW_NOTE =
	'Table row "keep together" (cantSplit) and repeating header rows (tblHeader) are not yet represented in the document model; every row may split across a page.';
const NEXT_COLUMN_NOTE = 'A "next column" section break is laid out as a continuous section break.';

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

/** Floating pictures anchored in a paragraph, with their `wp:positionH`/`wp:positionV` placement. */
function floatsOf(paragraph: Paragraph): { floats?: LayoutFloat[] } {
	const floats: LayoutFloat[] = [];
	for (const run of paragraph.runs) {
		const image = run.image;
		if (!image?.anchored || !image.partName) continue;
		const placement = image.placement;
		floats.push({
			partName: image.partName,
			contentType: image.contentType,
			widthPx: image.widthPx,
			heightPx: image.heightPx,
			...(placement?.relativeFrom ? { relativeFromH: placement.relativeFrom } : {}),
			...(placement?.align ? { alignH: placement.align } : {}),
			...(placement?.offsetXPx !== undefined ? { offsetXPx: placement.offsetXPx } : {}),
			...(placement?.relativeFromV ? { relativeFromV: placement.relativeFromV } : {}),
			...(placement?.alignV ? { alignV: placement.alignV } : {}),
			...(placement?.offsetYPx !== undefined ? { offsetYPx: placement.offsetYPx } : {}),
			...(placement?.behindText ? { behindText: true } : {}),
			...(placement?.wrap ? { wrap: placement.wrap } : {}),
		});
	}
	return floats.length ? { floats } : {};
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

	const now = new Date();
	const theme = model.theme;
	/** Effective run formatting: document defaults, paragraph and character styles, theme fonts. */
	function effective(run: TextRun, paragraphStyleId: string | undefined) {
		const formatting = resolveRunFormatting(run, {
			runCatalog: model.characterStyles,
			paragraphCatalog: catalog,
			paragraphStyleId,
		});
		const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
		const family = formatting.fontFamily ?? (role && theme ? theme.fonts[role]?.latin : undefined);
		const color =
			formatting.color ??
			(formatting.colorTheme && theme
				? resolveThemeColorReference(formatting.colorTheme, theme)
				: undefined);
		return { formatting, family, color };
	}
	function adaptRun(run: TextRun, paragraphStyleId: string | undefined): LayoutRun {
		const { formatting, family, color } = effective(run, paragraphStyleId);
		// DATE and TIME update when Word paginates for display or printing.
		const name = run.field ? fieldName(run.field.instr) : '';
		const text =
			run.field && (name === 'DATE' || name === 'TIME')
				? dateFieldResult(name, run.field.instr, now)
				: run.text;
		const image = run.image;
		const object =
			image && !image.anchored
				? {
						partName: image.partName,
						contentType: image.contentType,
						widthPx: image.widthPx,
						heightPx: image.heightPx,
					}
				: undefined;
		return {
			// Hidden text (`w:vanish`) takes no space when printed, as in Word by default.
			text: image || formatting.vanish ? '' : formatting.caps ? text.toUpperCase() : text,
			...(object ? { object } : {}),
			bold: formatting.bold,
			italic: formatting.italic,
			fontFamily: family,
			fontSizePt: formatting.fontSize,
			...(color && /^#[0-9a-f]{6}$/i.test(color) ? { color } : {}),
			...(formatting.underline ? { underline: true } : {}),
			...(formatting.strike || formatting.doubleStrike ? { strike: true } : {}),
			...(run.break ? { breakAfter: run.break } : {}),
		};
	}
	const labels = computeListLabels(model);
	/** The list label run (number or bullet plus its suffix), formatted like the paragraph's text. */
	function labelRun(paragraph: Paragraph, label: ParagraphListLabel): LayoutRun {
		const first = adaptRun(paragraph.runs[0] ?? { text: '' }, paragraph.style);
		const suffix = label.suffix === 'tab' ? '\t' : label.suffix === 'space' ? ' ' : '';
		return {
			text: `${displayListLabel(label.text)}${suffix}`,
			synthetic: true,
			bold: first.bold,
			italic: first.italic,
			fontFamily: first.fontFamily,
			fontSizePt: first.fontSizePt,
			...(first.color ? { color: first.color } : {}),
		};
	}
	function adaptParagraph(paragraph: Paragraph): LayoutParagraph {
		const resolved = catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph;
		reportOnce(KEEP_TOGETHER_NOTE);
		const label = labels.get(paragraph.id);
		// Numbering level indents apply unless the paragraph or its style sets its own.
		const ownFirstLine =
			resolved.firstLineTwips !== undefined || resolved.hangingTwips !== undefined;
		const runs = paragraph.runs.map((run) => adaptRun(run, paragraph.style));
		return {
			kind: 'paragraph',
			id: paragraph.id,
			runs: label ? [labelRun(paragraph, label), ...runs] : runs,
			...(paragraph.tabStops?.length
				? {
						tabStops: paragraph.tabStops.map((stop) => ({
							posPx: twipsToPx(stop.posTwips),
							align: stop.align,
							...(stop.leader ? { leader: stop.leader } : {}),
						})),
					}
				: {}),
			...floatsOf(paragraph),
			align: resolved.align,
			direction: resolved.direction,
			spacingBeforeTwips: resolved.spacingBeforeTwips,
			spacingAfterTwips: resolved.spacingAfterTwips,
			lineSpacingTwips: resolved.lineSpacingTwips,
			lineSpacingRule: resolved.lineSpacingRule,
			indentLeftTwips: resolved.indentLeftTwips ?? label?.indentLeftTwips,
			indentRightTwips: resolved.indentRightTwips,
			indentStartTwips: resolved.indentStartTwips,
			indentEndTwips: resolved.indentEndTwips,
			firstLineTwips: ownFirstLine ? resolved.firstLineTwips : label?.firstLineTwips,
			hangingTwips: ownFirstLine ? resolved.hangingTwips : label?.hangingTwips,
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
			...(section.verticalAlign ? { verticalAlign: section.verticalAlign } : {}),
			blocks: slice,
		};
	});
	return { sections: result };
}
