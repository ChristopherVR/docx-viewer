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
import { definedProps } from './defined-props.js';
import { adaptTable } from './adapt-table.js';
import { groupParagraphBorders, paragraphBox } from './adapt-paragraph-box.js';
import { endnoteParagraphs, noteLabels, paragraphFootnotes } from './adapt-notes.js';
import type {
	LayoutBlock,
	LayoutDocumentInput,
	LayoutFloat,
	LayoutParagraph,
	LayoutParagraphBorders,
	LayoutRun,
	LayoutSection,
	LayoutTable,
} from './input.js';

const TABLE_ROW_NOTE =
	'Table rows with an exact height clip content that does not fit; text in them is not shrunk.';
const VERTICAL_MERGE_NOTE =
	'Vertically merged table cells are drawn as one cell, but their text stays in the first row of the merge.';
const NEXT_COLUMN_NOTE = 'A "next column" section break is laid out as a continuous section break.';

const twipsToPx = (twips: number): number => twips / 15;

function sectionBreak(
	type: SectionProperties['type'],
	note: (message: string) => void,
): NonNullable<LayoutSection['break']> | undefined {
	if (type === 'nextColumn') {
		note(NEXT_COLUMN_NOTE);
		return 'continuous';
	}
	return type;
}

/** Floating pictures anchored in a paragraph, with their `wp:positionH`/`wp:positionV` placement. */
export function paragraphFloats(paragraph: Paragraph): LayoutFloat[] {
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
	return floats;
}
function floatsOf(paragraph: Paragraph): { floats?: LayoutFloat[] } {
	const floats = paragraphFloats(paragraph);
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
	const noteLabel = noteLabels(model);
	function adaptRun(
		run: TextRun,
		paragraphStyleId: string | undefined,
		markLabel?: string,
	): LayoutRun {
		const { formatting, family, color } = effective(run, paragraphStyleId);
		// DATE and TIME update when Word paginates for display or printing.
		const name = run.field ? fieldName(run.field.instr) : '';
		// Note references and a note's own number mark show the note's number, raised.
		const noteText = run.noteReference
			? noteLabel(run.noteReference.kind, run.noteReference.id)
			: run.noteMark
				? (markLabel ?? '')
				: undefined;
		const text =
			noteText ??
			(run.field && (name === 'DATE' || name === 'TIME')
				? dateFieldResult(name, run.field.instr, now)
				: run.text);
		const script =
			noteText !== undefined || formatting.verticalAlign === 'superscript'
				? 'super'
				: formatting.verticalAlign === 'subscript'
					? 'sub'
					: undefined;
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
			...definedProps({
				bold: formatting.bold,
				italic: formatting.italic,
				fontFamily: family,
				fontSizePt: formatting.fontSize,
			}),
			...(color && /^#[0-9a-f]{6}$/i.test(color) ? { color } : {}),
			...(formatting.underline ? { underline: true } : {}),
			...(formatting.strike || formatting.doubleStrike ? { strike: true } : {}),
			...(run.break ? { breakAfter: run.break } : {}),
			...(script ? { script } : {}),
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
			...definedProps({
				bold: first.bold,
				italic: first.italic,
				fontFamily: first.fontFamily,
				fontSizePt: first.fontSizePt,
			}),
			...(first.color ? { color: first.color } : {}),
		};
	}
	const betweenBorders = new Map<LayoutParagraph, LayoutParagraphBorders['top']>();
	function adaptParagraph(paragraph: Paragraph, markLabel?: string): LayoutParagraph {
		const adapted = adaptParagraphOnly(paragraph, markLabel);
		const between = paragraphBox(
			catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph,
			model.theme,
		).betweenBorder;
		if (between) betweenBorders.set(adapted, between);
		return adapted;
	}
	function adaptParagraphOnly(paragraph: Paragraph, markLabel?: string): LayoutParagraph {
		const resolved = catalog ? resolveParagraphFormatting(paragraph, catalog) : paragraph;
		const label = labels.get(paragraph.id);
		// Numbering level indents apply unless the paragraph or its style sets its own.
		const ownFirstLine =
			resolved.firstLineTwips !== undefined || resolved.hangingTwips !== undefined;
		const runs = paragraph.runs.map((run) => adaptRun(run, paragraph.style, markLabel));
		const footnotes = paragraphFootnotes(paragraph, model, noteLabel, adaptParagraph);
		const { betweenBorder: _between, ...box } = paragraphBox(resolved, model.theme);
		return {
			kind: 'paragraph',
			id: paragraph.id,
			runs: label ? [labelRun(paragraph, label), ...runs] : runs,
			...(footnotes.length ? { footnotes } : {}),
			...box,
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
			...definedProps({
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
			}),
			...(paragraph.pageBreakBefore ? { pageBreakBefore: true } : {}),
			...(resolved.keepNext ? { keepNext: true } : {}),
			...(resolved.keepLines ? { keepLines: true } : {}),
			...(resolved.widowControl === false ? { widowControl: false } : {}),
			...(resolved.contextualSpacing ? { contextualSpacing: true } : {}),
		};
	}
	function adaptBlock(block: Block): LayoutBlock {
		if (block.type === 'paragraph') return adaptParagraph(block);
		reportOnce(TABLE_ROW_NOTE);
		reportOnce(VERTICAL_MERGE_NOTE);
		return adaptTable(block, model, adaptParagraph);
	}

	const blocks = [
		...model.blocks.map(adaptBlock),
		...endnoteParagraphs(model, noteLabel, adaptParagraph),
	];
	groupParagraphBorders(blocks, betweenBorders);
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
		const sectionStart = index === 0 ? undefined : sectionBreak(section.type, reportOnce);
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
			...(sectionStart && { break: sectionStart }),
			...(section.verticalAlign ? { verticalAlign: section.verticalAlign } : {}),
			blocks: slice,
		};
	});
	return { sections: result };
}
