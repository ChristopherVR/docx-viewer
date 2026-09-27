import { EditorState, Transaction } from 'prosemirror-state';
import type { DocumentModel, Block, Paragraph, Table, TextRun } from '@christophervr/docx-core';
import { computeListLabels, numberNotesInOrder } from '@christophervr/docx-core';
import { schema } from './schema';
import { extraRunProperties } from './run-extra-mark';
import { appendInlineNode, runToInlineNodes, type NoteNumberLookup } from './run-adapter';
import { convertMergedTable, convertSimpleTable, tableNode } from './table-model-adapter';

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
			listLabelText: label?.text ?? null,
			listSuffix: label?.suffix ?? null,
			listIndentLeftTwips: label?.indentLeftTwips ?? null,
			listHangingTwips: label?.hangingTwips ?? null,
			listFirstLineTwips: label?.firstLineTwips ?? null,
			pageBreakBefore: paragraph.pageBreakBefore ?? false,
		},
		children,
	);
}

export function modelToDoc(model: DocumentModel) {
	const labels = computeListLabels(model);
	const footnoteOrder = numberNotesInOrder(model.blocks, 'footnote');
	const endnoteOrder = numberNotesInOrder(model.blocks, 'endnote');
	const noteNumber: NoteNumberLookup = (kind, id) =>
		(kind === 'footnote' ? footnoteOrder : endnoteOrder).get(id) ?? 1;
	const blocks = model.blocks.map((block) => {
		if (block.type === 'paragraph') return paragraphNode(block, labels, noteNumber);
		return tableNode(block, (paragraph) => paragraphNode(paragraph, labels, noteNumber));
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
		},
		blocks,
	);
}

function sameRuns(left: TextRun[], right: TextRun[]) {
	const compact = (runs: TextRun[]) =>
		runs.reduce<TextRun[]>((result, run) => {
			const normalized = {
				...run,
				bold: run.bold || undefined,
				italic: run.italic || undefined,
				underline: run.underline || undefined,
				strike: run.strike || undefined,
			};
			const previous = result.at(-1);
			const fields: (keyof TextRun)[] = [
				'bold',
				'italic',
				'underline',
				'strike',
				'highlight',
				'verticalAlign',
				'language',
				'eastAsiaLanguage',
				'bidiLanguage',
				'rtl',
				'fontFamily',
				'fontSize',
				'color',
				'break',
			];
			if (
				previous &&
				!previous.noteReference &&
				!normalized.noteReference &&
				fields.every((field) => previous[field] === normalized[field]) &&
				sameRevisionMeta(previous, normalized)
			)
				previous.text += normalized.text;
			else result.push(normalized);
			return result;
		}, []);
	const a = compact(left);
	const b = compact(right);
	const fields: (keyof TextRun)[] = [
		'bold',
		'italic',
		'underline',
		'strike',
		'highlight',
		'verticalAlign',
		'language',
		'eastAsiaLanguage',
		'bidiLanguage',
		'rtl',
		'fontFamily',
		'fontSize',
		'color',
		'break',
	];
	const sameNoteReference = (x?: TextRun['noteReference'], y?: TextRun['noteReference']) =>
		x?.kind === y?.kind && x?.id === y?.id;
	return (
		a.length === b.length &&
		a.every(
			(run, index) =>
				run.text === b[index].text &&
				fields.every((field) => run[field] === b[index][field]) &&
				sameNoteReference(run.noteReference, b[index].noteReference) &&
				sameRevisionMeta(run, b[index]),
		)
	);
}
/** `revision`/`commentIds` are objects/arrays; compare by value rather than by reference. */
function sameRevisionMeta(left: TextRun, right: TextRun): boolean {
	return (
		JSON.stringify(left.revision) === JSON.stringify(right.revision) &&
		JSON.stringify(left.commentIds) === JSON.stringify(right.commentIds) &&
		JSON.stringify(extraRunProperties(left)) === JSON.stringify(extraRunProperties(right))
	);
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
			Boolean(previous.pageBreakBefore) === Boolean(node.attrs.pageBreakBefore)
		)
			return previous;
		return {
			type: 'paragraph',
			id,
			runs,
			...(node.attrs.align != null ? { align: node.attrs.align } : {}),
			...(node.attrs.direction != null ? { direction: node.attrs.direction } : {}),
			...(node.attrs.style ? { style: node.attrs.style } : {}),
			...(node.attrs.spacingBeforeTwips != null
				? { spacingBeforeTwips: node.attrs.spacingBeforeTwips }
				: {}),
			...(node.attrs.spacingAfterTwips != null
				? { spacingAfterTwips: node.attrs.spacingAfterTwips }
				: {}),
			...(node.attrs.lineSpacingTwips != null
				? { lineSpacingTwips: node.attrs.lineSpacingTwips }
				: {}),
			...(node.attrs.lineSpacingRule != null
				? { lineSpacingRule: node.attrs.lineSpacingRule }
				: {}),
			...(node.attrs.indentLeftTwips != null
				? { indentLeftTwips: node.attrs.indentLeftTwips }
				: {}),
			...(node.attrs.indentRightTwips != null
				? { indentRightTwips: node.attrs.indentRightTwips }
				: {}),
			...(node.attrs.indentStartTwips != null
				? { indentStartTwips: node.attrs.indentStartTwips }
				: {}),
			...(node.attrs.indentEndTwips != null ? { indentEndTwips: node.attrs.indentEndTwips } : {}),
			...(node.attrs.firstLineTwips != null ? { firstLineTwips: node.attrs.firstLineTwips } : {}),
			...(node.attrs.hangingTwips != null ? { hangingTwips: node.attrs.hangingTwips } : {}),
			...(node.attrs.numId != null
				? { numbering: { numId: node.attrs.numId, level: node.attrs.ilvl ?? 0 } }
				: {}),
			...(node.attrs.pageBreakBefore ? { pageBreakBefore: true } : {}),
		};
	};

	const blocks: Block[] = [];
	doc.forEach((node) => {
		if (node.type.name === 'paragraph') blocks.push(convertParagraph(node as typeof doc));
		if (node.type.name === 'table') {
			const id = String(node.attrs.id || `t-edit-${++nextId}`);
			const prior = priorTables.get(id);
			const asParagraph = (n: unknown) => convertParagraph(n as typeof doc);
			if (prior && prior.structureEditable === false) blocks.push(convertMergedTable(node, prior, asParagraph));
			else
				blocks.push({
					// Table-level properties (grid, width, borders, style, look) have no editor node;
					// carry them over from the prior model so editing cell text never drops them.
					...(prior ? { ...prior, structureEditable: undefined } : {}),
					type: 'table',
					id,
					rows: convertSimpleTable(node, asParagraph),
					...(node.attrs.structureEditable === false ? { structureEditable: false } : {}),
				} as Table);
		}
	});
	return {
		...prior,
		blocks,
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
