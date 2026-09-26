import { EditorState, Transaction } from 'prosemirror-state';
import type { DocumentModel, Block, Paragraph, Table, TextRun } from '@christophervr/docx-core';
import { schema } from './schema';

function paragraphNode(paragraph: Paragraph) {
	const children = paragraph.runs
		.filter((run) => run.text.length)
		.map((run) => {
			const marks = [];
			if (run.bold) marks.push(schema.marks.bold.create());
			if (run.italic) marks.push(schema.marks.italic.create());
			if (run.underline) marks.push(schema.marks.underline.create());
			if (run.strike) marks.push(schema.marks.strike.create());
			if (run.highlight) marks.push(schema.marks.highlight.create({ color: run.highlight }));
			if (run.verticalAlign)
				marks.push(schema.marks.verticalAlign.create({ value: run.verticalAlign }));
			if (run.fontFamily || run.fontSize || run.color) {
				marks.push(
					schema.marks.font.create({
						family: run.fontFamily || null,
						size: run.fontSize || null,
						color: run.color || null,
					}),
				);
			}
			return schema.text(run.text, marks);
		});
	return schema.node(
		'paragraph',
		{
			id: paragraph.id,
			align: paragraph.align || 'left',
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
		const rows = block.rows.map((row) =>
			schema.node(
				'tableRow',
				null,
				row.map((cell) => schema.node('tableCell', null, cell.paragraphs.map(paragraphNode))),
			),
		);
		return schema.node(
			'table',
			{
				id: block.id,
				structureEditable:
					(block as Table & { structureEditable?: boolean }).structureEditable !== false,
			},
			rows,
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
	const remember = (paragraph: Paragraph) => previousParagraphs.set(paragraph.id, paragraph);
	for (const block of prior.blocks) {
		if (block.type === 'paragraph') remember(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(remember);
	}

	const convertParagraph = (node: typeof doc): Paragraph => {
		const runs: TextRun[] = [];
		node.forEach((child) => {
			if (!child.isText) return;
			const run: TextRun = { text: child.text || '' };
			if (child.marks.some((mark) => mark.type.name === 'bold')) run.bold = true;
			if (child.marks.some((mark) => mark.type.name === 'italic')) run.italic = true;
			if (child.marks.some((mark) => mark.type.name === 'underline')) run.underline = true;
			if (child.marks.some((mark) => mark.type.name === 'strike')) run.strike = true;
			const highlight = child.marks.find((mark) => mark.type.name === 'highlight');
			if (highlight?.attrs.color) run.highlight = highlight.attrs.color;
			const verticalAlign = child.marks.find((mark) => mark.type.name === 'verticalAlign');
			if (verticalAlign?.attrs.value) run.verticalAlign = verticalAlign.attrs.value;
			const font = child.marks.find((mark) => mark.type.name === 'font');
			if (font?.attrs.family) run.fontFamily = font.attrs.family;
			if (font?.attrs.size) run.fontSize = font.attrs.size;
			if (font?.attrs.color) run.color = font.attrs.color;
			const previousRun = runs.at(-1);
			const fields: (keyof TextRun)[] = [
				'bold',
				'italic',
				'underline',
				'strike',
				'highlight',
				'verticalAlign',
				'fontFamily',
				'fontSize',
				'color',
			];
			if (previousRun && fields.every((field) => previousRun[field] === run[field]))
				previousRun.text += run.text;
			else runs.push(run);
		});
		if (!runs.length) runs.push({ text: '' });
		const id = String(node.attrs.id || `p-edit-${++nextId}`);
		const previous = previousParagraphs.get(id);
		if (
			previous &&
			sameRuns(previous.runs, runs) &&
			(previous.align || 'left') === (node.attrs.align || 'left') &&
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
			...(node.attrs.align === 'left' && previous?.align == null
				? {}
				: { align: node.attrs.align }),
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
			...(node.attrs.lineSpacingTwips != null && node.attrs.lineSpacingRule
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
			const rows: Table['rows'] = [];
			node.forEach((row) => {
				const cells: Table['rows'][number] = [];
				row.forEach((cell) => {
					const paragraphs: Paragraph[] = [];
					cell.forEach((paragraph) => paragraphs.push(convertParagraph(paragraph as typeof doc)));
					cells.push({ paragraphs });
				});
				rows.push(cells);
			});
			blocks.push({
				type: 'table',
				id: String(node.attrs.id || `t-edit-${++nextId}`),
				rows,
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
