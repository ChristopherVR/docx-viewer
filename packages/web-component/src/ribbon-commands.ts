import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { DEFAULT_TABLE_BORDERS } from '@christophervr/docx-core';
import { schema } from './schema';
import { parseLineSpacingValue } from './line-spacing';

export function applyFont(view: EditorView, key: 'family' | 'size' | 'color', value: string) {
	const attr = key === 'family' ? 'family' : key === 'size' ? 'size' : 'color';
	const mark = schema.marks.font;
	const nextValue = key === 'size' ? Number(value) : value;
	const { state } = view;
	if (state.selection.empty) {
		const marks = state.storedMarks || state.selection.$from.marks();
		const existing = marks.find((item) => item.type === mark);
		const type = mark.create({ ...existing?.attrs, [attr]: nextValue });
		view.dispatch(state.tr.setStoredMarks([...marks.filter((item) => item.type !== mark), type]));
		return;
	}
	let transaction = state.tr;
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (!node.isText) return;
		const start = Math.max(pos, state.selection.from);
		const end = Math.min(pos + node.nodeSize, state.selection.to);
		const existing = node.marks.find((item) => item.type === mark);
		const type = mark.create({ ...existing?.attrs, [attr]: nextValue });
		transaction = transaction.removeMark(start, end, mark).addMark(start, end, type);
	});
	view.dispatch(transaction);
}

export function clearFormatting(view: EditorView) {
	const { state } = view;
	let transaction = state.tr;
	for (const mark of Object.values(schema.marks))
		transaction = transaction.removeMark(state.selection.from, state.selection.to, mark);
	view.dispatch(transaction.setStoredMarks([]));
}

let tableId = 0;
export function insertTable(view: EditorView, idGenerator?: (kind: string) => string) {
	const ids = new Set<string>();
	view.state.doc.descendants((node) => {
		if (typeof node.attrs.id === 'string' && node.attrs.id) ids.add(node.attrs.id);
	});
	const nextId = (kind: string) => {
		let candidate = '';
		do candidate = idGenerator ? idGenerator(kind) : `dve-${kind}-${++tableId}`;
		while (ids.has(candidate));
		ids.add(candidate);
		return candidate;
	};
	const paragraph = () => schema.nodes.paragraph.create({ id: nextId('cell') });
	const table = schema.nodes.table.create(
		{ id: nextId('table'), borders: JSON.stringify(DEFAULT_TABLE_BORDERS) },
		[0, 1].map(() =>
			schema.nodes.tableRow.create(
				null,
				[0, 1].map(() => schema.nodes.tableCell.create(null, paragraph())),
			),
		),
	);
	view.dispatch(closeHistory(view.state.tr.replaceSelectionWith(table)).scrollIntoView());
}

export function updatePage(view: EditorView, key: 'margin' | 'orientation', value: string) {
	const page = view.state.doc.attrs;
	if (key === 'margin') {
		const margin = value === 'narrow' ? 48 : value === 'wide' ? 144 : 96;
		view.dispatch(
			view.state.tr
				.setDocAttribute('marginTop', margin)
				.setDocAttribute('marginRight', margin)
				.setDocAttribute('marginBottom', margin)
				.setDocAttribute('marginLeft', margin),
		);
		return;
	}
	const landscape = value === 'landscape';
	const long = Math.max(page.pageWidth, page.pageHeight);
	const short = Math.min(page.pageWidth, page.pageHeight);
	view.dispatch(
		view.state.tr
			.setDocAttribute('pageWidth', landscape ? long : short)
			.setDocAttribute('pageHeight', landscape ? short : long),
	);
}

export function updateParagraphs(
	view: EditorView,
	key: 'indent' | 'spacingBefore' | 'spacingAfter' | 'lineSpacing',
	value: string,
) {
	const { state } = view;
	if (key === 'lineSpacing') {
		if (value !== 'inherit' && !parseLineSpacingValue(value)) return;
	} else if (
		key !== 'indent' &&
		value !== 'inherit' &&
		(!Number.isFinite(Number(value)) || Number(value) < 0)
	)
		return;
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	let transaction = state.tr;
	for (const pos of positions) {
		const node = transaction.doc.nodeAt(pos);
		if (!node) continue;
		if (key === 'indent') {
			const logical = node.attrs.indentStartTwips != null;
			const attr = logical ? 'indentStartTwips' : 'indentLeftTwips';
			const current = Number(node.attrs[attr] ?? 0);
			const next = Math.max(0, current + (value === 'increase' ? 360 : -360));
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				[attr]: next,
			});
		} else if (key === 'spacingAfter') {
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				spacingAfterTwips: value === 'inherit' ? null : Number(value),
			});
		} else if (key === 'spacingBefore') {
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				spacingBeforeTwips: value === 'inherit' ? null : Number(value),
			});
		} else {
			const lineSpacing = value === 'inherit' ? null : parseLineSpacingValue(value);
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				lineSpacingTwips: lineSpacing?.twips ?? null,
				lineSpacingRule: lineSpacing?.rule ?? null,
			});
		}
	}
	if (transaction.docChanged) view.dispatch(closeHistory(transaction).scrollIntoView());
}

export function applyPageStyles(paper: HTMLElement, model: DocumentModel, zoom: number) {
	const page = model.page;
	paper.style.width = `${page.width}px`;
	paper.style.minHeight = `${page.height}px`;
	paper.style.padding = `${page.marginTop}px ${page.marginRight}px ${page.marginBottom}px ${page.marginLeft}px`;
	paper.style.transform = 'none';
	paper.style.setProperty('--dve-zoom', String(zoom));
}
