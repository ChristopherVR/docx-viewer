import { EditorState, Transaction } from 'prosemirror-state';
import type { DocumentModel, Block, Paragraph, Table, TextRun } from '@christophervr/docx-core';
import { schema } from './schema';
import { appendInlineNode, runToInlineNodes } from './run-adapter';
import { tableNode, convertSimpleTable, convertMergedTable } from './table-model-adapter';

function paragraphNode(paragraph: Paragraph) {
	const children = paragraph.runs.flatMap(runToInlineNodes);
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
		},
		children,
	);
}

export function modelToDoc(model: DocumentModel) {
	const blocks = model.blocks.map((block) => {
		if (block.type === 'paragraph') return paragraphNode(block);
		return tableNode(block, paragraphNode);
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
			];
			if (previous && fields.every((field) => previous[field] === normalized[field]))
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
	];
	return (
		a.length === b.length &&
		a.every(
			(run, index) =>
				run.text === b[index].text && fields.every((field) => run[field] === b[index][field]),
		)
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
			previous.hangingTwips === (node.attrs.hangingTwips ?? undefined)
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
