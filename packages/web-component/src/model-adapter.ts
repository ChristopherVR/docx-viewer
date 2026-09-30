import { EditorState, Transaction } from 'prosemirror-state';
import type {
	DocumentModel,
	Block,
	Paragraph,
	SectionProperties,
	Table,
	TextRun,
} from '@christophervr/docx-core';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	computeListLabels,
	displayListLabel,
	formatNoteNumber,
	numberNotesInOrder,
} from '@christophervr/docx-core';
import { schema } from './schema';
import { paragraphTwipsFromAttrs } from './attr-units';
import { sameJson, sameRuns } from './run-compare';
import { appendInlineNode, runToInlineNodes, type NoteNumberLookup } from './run-adapter';
import { convertMergedTable, convertSimpleTable, tableNode } from './table-model-adapter';
import { parseBordersJson } from './table-render';
import { sectionLayoutJson, sectionsFromLayout } from './section-layout';
import { sectionsOf } from './section-commands';

type ListLabels = ReturnType<typeof computeListLabels>;

function paragraphNode(paragraph: Paragraph, labels: ListLabels, noteNumber?: NoteNumberLookup) {
	const label = labels.get(paragraph.id);
	const children = paragraph.runs.flatMap((run) => runToInlineNodes(run, noteNumber));
	return schema.node(
		'paragraph',
		{
			id: paragraph.id,
			align: paragraph.align ?? null,
			direction: paragraph.direction ?? null,
			style: paragraph.style || '',
			spacingBeforeTwips: paragraph.spacingBeforeTwips ?? null,
			spacingAfterTwips: paragraph.spacingAfterTwips ?? null,
			lineSpacingTwips: paragraph.lineSpacingTwips ?? null,
			lineSpacingRule: paragraph.lineSpacingRule ?? null,
			indentLeftTwips: paragraph.indentLeftTwips ?? null,
			indentRightTwips: paragraph.indentRightTwips ?? null,
			indentStartTwips: paragraph.indentStartTwips ?? null,
			indentEndTwips: paragraph.indentEndTwips ?? null,
			firstLineTwips: paragraph.firstLineTwips ?? null,
			hangingTwips: paragraph.hangingTwips ?? null,
			numId: paragraph.numbering?.numId ?? null,
			ilvl: paragraph.numbering ? paragraph.numbering.level : null,
			listLabelText: label ? displayListLabel(label.text) : null,
			listSuffix: label?.suffix ?? null,
			listIndentLeftTwips: label?.indentLeftTwips ?? null,
			listHangingTwips: label?.hangingTwips ?? null,
			listFirstLineTwips: label?.firstLineTwips ?? null,
			pageBreakBefore: paragraph.pageBreakBefore ?? false,
			tabStops: paragraph.tabStops?.length ? paragraph.tabStops : null,
			keepNext: paragraph.keepNext ?? null,
			keepLines: paragraph.keepLines ?? null,
			widowControl: paragraph.widowControl ?? null,
			contextualSpacing: paragraph.contextualSpacing ?? null,
			dropCap: paragraph.dropCap ?? null,
			borders: paragraph.borders ?? null,
			shadingFill: paragraph.shadingFill ?? null,
			bookmarks: paragraph.bookmarks ?? [],
		},
		children,
	);
}

/** Paragraph pagination toggles carried through the editor as node attributes. */
const KEEP_KEYS = ['keepNext', 'keepLines', 'widowControl', 'contextualSpacing'] as const;

export function modelToDoc(model: DocumentModel) {
	const labels = computeListLabels(model);
	const footnoteOrder = numberNotesInOrder(model.blocks, 'footnote');
	const endnoteOrder = numberNotesInOrder(model.blocks, 'endnote');
	const noteNumber: NoteNumberLookup = (kind, id) => {
		const number = (kind === 'footnote' ? footnoteOrder : endnoteOrder).get(id) ?? 1;
		const format =
			kind === 'footnote'
				? (model.footnoteNumFmt ?? 'decimal')
				: (model.endnoteNumFmt ?? 'lowerRoman');
		return { number, label: formatNoteNumber(number, format) };
	};
	const blocks = model.blocks.map((block) => {
		if (block.type === 'paragraph') return paragraphNode(block, labels, noteNumber);
		return tableNode(
			block,
			(paragraph) => paragraphNode(paragraph, labels, noteNumber),
			model.tableStyles,
		);
	});
	return schema.node(
		'doc',
		{
			pageWidth: model.page.width,
			pageHeight: model.page.height,
			marginTop: model.page.marginTop,
			marginRight: model.page.marginRight,
			marginBottom: model.page.marginBottom,
			marginLeft: model.page.marginLeft,
			sections: model.sections ? sectionLayoutJson(model.sections) : null,
			evenAndOddHeaders: Boolean(model.evenAndOddHeaders),
			pageColor: model.pageColor ?? null,
			autoHyphenation: Boolean(model.autoHyphenation),
			footnoteNumFmt: model.footnoteNumFmt ?? null,
			endnoteNumFmt: model.endnoteNumFmt ?? null,
		},
		blocks,
	);
}

/** A table inserted in the editor keeps the borders it was created with (see insertTable). */
function tableBordersFromNode(value: unknown): Partial<Table> {
	const borders = parseBordersJson(value);
	return borders ? { borders: borders as NonNullable<Table['borders']> } : {};
}
/**
 * A document with no recorded section layout (a new document nobody changed page setup on) has no
 * sections of its own, but header/footer content added since must survive: one section is
 * synthesized from the page geometry and keeps that content.
 */
function keepHeaderFooterSections(
	prior: SectionProperties[] | undefined,
	blocks: Block[],
	doc: ProseMirrorNode,
): { sections?: SectionProperties[] } {
	const source = prior?.find((section) => section.headers || section.footers);
	if (!source) return {};
	const page = {
		width: doc.attrs.pageWidth,
		height: doc.attrs.pageHeight,
		marginTop: doc.attrs.marginTop,
		marginRight: doc.attrs.marginRight,
		marginBottom: doc.attrs.marginBottom,
		marginLeft: doc.attrs.marginLeft,
	};
	const [section] = sectionsOf({ blocks, page } as DocumentModel);
	if (!section) return {};
	return {
		sections: [
			{
				...section,
				...(source.headers ? { headers: source.headers } : {}),
				...(source.footers ? { footers: source.footers } : {}),
			},
		],
	};
}

export function docToModel(
	doc: ReturnType<typeof modelToDoc>,
	prior: DocumentModel,
): DocumentModel {
	let nextId = 0;
	const previousParagraphs = new Map<string, Paragraph>();
	const priorTables = new Map<string, Table>();
	const remember = (paragraph: Paragraph) => previousParagraphs.set(paragraph.id, paragraph);
	for (const block of prior.blocks) {
		if (block.type === 'paragraph') remember(block);
		else {
			priorTables.set(block.id, block);
			for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(remember);
		}
	}

	const convertParagraph = (node: typeof doc): Paragraph => {
		const runs: TextRun[] = [];
		node.forEach((child) => appendInlineNode(runs, child));
		if (!runs.length) runs.push({ text: '' });
		const id = String(node.attrs.id || `p-edit-${++nextId}`);
		const previous = previousParagraphs.get(id);
		if (
			previous &&
			sameRuns(previous.runs, runs) &&
			previous.align === (node.attrs.align ?? undefined) &&
			previous.direction === (node.attrs.direction ?? undefined) &&
			(previous.style || '') === (node.attrs.style || '') &&
			previous.spacingBeforeTwips === (node.attrs.spacingBeforeTwips ?? undefined) &&
			previous.spacingAfterTwips === (node.attrs.spacingAfterTwips ?? undefined) &&
			previous.lineSpacingTwips === (node.attrs.lineSpacingTwips ?? undefined) &&
			previous.lineSpacingRule === (node.attrs.lineSpacingRule ?? undefined) &&
			previous.indentLeftTwips === (node.attrs.indentLeftTwips ?? undefined) &&
			previous.indentRightTwips === (node.attrs.indentRightTwips ?? undefined) &&
			previous.indentStartTwips === (node.attrs.indentStartTwips ?? undefined) &&
			previous.indentEndTwips === (node.attrs.indentEndTwips ?? undefined) &&
			previous.firstLineTwips === (node.attrs.firstLineTwips ?? undefined) &&
			previous.hangingTwips === (node.attrs.hangingTwips ?? undefined) &&
			(previous.numbering?.numId ?? null) === (node.attrs.numId ?? null) &&
			(previous.numbering ? previous.numbering.level : null) === (node.attrs.ilvl ?? null) &&
			Boolean(previous.pageBreakBefore) === Boolean(node.attrs.pageBreakBefore) &&
			sameJson(previous.tabStops ?? null, node.attrs.tabStops ?? null) &&
			KEEP_KEYS.every((key) => (previous[key] ?? null) === (node.attrs[key] ?? null)) &&
			sameJson(previous.dropCap ?? null, node.attrs.dropCap ?? null) &&
			sameJson(previous.borders ?? null, node.attrs.borders ?? null) &&
			(previous.shadingFill ?? null) === (node.attrs.shadingFill ?? null) &&
			sameJson(previous.bookmarks ?? [], node.attrs.bookmarks ?? [])
		)
			return previous;
		return {
			type: 'paragraph',
			id,
			runs,
			...(node.attrs.align != null ? { align: node.attrs.align } : {}),
			...(node.attrs.direction != null ? { direction: node.attrs.direction } : {}),
			...(node.attrs.style ? { style: node.attrs.style } : {}),
			...paragraphTwipsFromAttrs(node.attrs),
			...(node.attrs.lineSpacingRule != null
				? { lineSpacingRule: node.attrs.lineSpacingRule }
				: {}),
			...(node.attrs.numId != null
				? { numbering: { numId: node.attrs.numId, level: node.attrs.ilvl ?? 0 } }
				: {}),
			...(node.attrs.pageBreakBefore ? { pageBreakBefore: true } : {}),
			...(node.attrs.tabStops?.length ? { tabStops: structuredClone(node.attrs.tabStops) } : {}),
			...Object.fromEntries(
				KEEP_KEYS.filter((key) => node.attrs[key] != null).map((key) => [key, node.attrs[key]]),
			),
			...(node.attrs.dropCap ? { dropCap: { ...node.attrs.dropCap } } : {}),
			...(node.attrs.borders ? { borders: structuredClone(node.attrs.borders) } : {}),
			...(node.attrs.shadingFill ? { shadingFill: node.attrs.shadingFill } : {}),
			...(node.attrs.bookmarks?.length ? { bookmarks: [...node.attrs.bookmarks] } : {}),
		};
	};

	const blocks: Block[] = [];
	doc.forEach((node) => {
		if (node.type.name === 'paragraph') blocks.push(convertParagraph(node as typeof doc));
		if (node.type.name === 'table') {
			const id = String(node.attrs.id || `t-edit-${++nextId}`);
			const prior = priorTables.get(id);
			const asParagraph = (n: unknown) => convertParagraph(n as typeof doc);
			if (prior && prior.structureEditable === false)
				blocks.push(convertMergedTable(node, prior, asParagraph));
			else
				blocks.push({
					// Table-level properties (grid, width, borders, style, look) have no editor node;
					// carry them over from the prior model so editing cell text never drops them.
					...(prior
						? {
								...prior,
								structureEditable: undefined,
								// Row properties follow rows by position; drop them once rows are added or removed.
								...(prior.rowProperties && prior.rows.length !== node.childCount
									? { rowProperties: undefined }
									: {}),
							}
						: tableBordersFromNode(node.attrs.borders)),
					type: 'table',
					id,
					rows: convertSimpleTable(node, asParagraph, prior),
					...(node.attrs.structureEditable === false ? { structureEditable: false } : {}),
				} as Table);
		}
	});
	// The editor document is the source of truth for section layout (so undo covers page setup).
	const {
		sections: priorSections,
		evenAndOddHeaders: _evenOdd,
		pageColor: _pageColor,
		autoHyphenation: _hyphenation,
		...rest
	} = prior;
	return {
		...rest,
		blocks,
		...(doc.attrs.evenAndOddHeaders ? { evenAndOddHeaders: true } : {}),
		...(doc.attrs.pageColor ? { pageColor: String(doc.attrs.pageColor) } : {}),
		...(doc.attrs.autoHyphenation ? { autoHyphenation: true } : {}),
		...(typeof doc.attrs.sections === 'string'
			? { sections: sectionsFromLayout(doc.attrs.sections, priorSections, blocks) }
			: keepHeaderFooterSections(priorSections, blocks, doc)),
		page: {
			width: doc.attrs.pageWidth,
			height: doc.attrs.pageHeight,
			marginTop: doc.attrs.marginTop,
			marginRight: doc.attrs.marginRight,
			marginBottom: doc.attrs.marginBottom,
			marginLeft: doc.attrs.marginLeft,
		},
	};
}

export function assignMissingParagraphIds(state: EditorState): Transaction | null {
	const reserved = new Set<string>();
	const seen = new Set<string>();
	state.doc.descendants((node) => {
		if (node.type.name === 'paragraph' && node.attrs.id) reserved.add(String(node.attrs.id));
	});
	let next = 1;
	let transaction = state.tr;
	state.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return;
		const id = String(node.attrs.id || '');
		if (id && !seen.has(id)) {
			seen.add(id);
			return;
		}
		let generated = '';
		do {
			generated = `p-edit-${next++}`;
		} while (reserved.has(generated));
		reserved.add(generated);
		seen.add(generated);
		transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, id: generated });
	});
	return transaction.docChanged ? transaction : null;
}
