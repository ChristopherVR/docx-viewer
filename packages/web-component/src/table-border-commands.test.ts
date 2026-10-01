// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Table } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import {
	applyCellBorderSettings,
	readCellBorderSettings,
	setCellBorders,
	tableBorderContext,
} from './table-border-commands';

function sample(): DocumentModel {
	const model = createDocument();
	const cell = (id: string) => ({
		paragraphs: [{ type: 'paragraph' as const, id, runs: [{ text: id }] }],
	});
	model.blocks = [
		{
			type: 'table',
			id: 't',
			rows: [
				[cell('a'), cell('b'), cell('c')],
				[cell('d'), cell('e'), cell('f')],
			],
		} as Table,
	];
	return model;
}
function viewAt(model: DocumentModel, paragraphId: string, toId = paragraphId) {
	const doc = modelToDoc(model);
	const find = (id: string) => {
		let found = 0;
		doc.descendants((node, pos) => {
			if (node.type.name === 'paragraph' && node.attrs.id === id) found = pos + 1;
		});
		return found;
	};
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc,
			selection: TextSelection.create(doc, find(paragraphId), find(toId)),
		}),
	});
}
const cells = (view: EditorView, model: DocumentModel) =>
	(docToModel(view.state.doc, model).blocks[0] as Table).rows;

describe('table cell borders', () => {
	it('finds the cells a text selection touches as one rectangle', () => {
		const view = viewAt(sample(), 'a', 'e');
		const ctx = tableBorderContext(view.state)!;
		expect(ctx.rect).toEqual({ top: 0, bottom: 1, left: 0, right: 2 });
		expect(ctx.cells).toHaveLength(5);
	});

	it('draws a side on the caret cell and the neighbour that shares the edge, then toggles it off', () => {
		const model = sample();
		const view = viewAt(model, 'b');
		expect(setCellBorders(view, 'right')).toBe(true);
		let rows = cells(view, model);
		expect(rows[0]![1]!.borders?.right?.style).toBe('single');
		expect(rows[0]![2]!.borders?.left?.style).toBe('single');
		expect(rows[0]![0]!.borders).toBeUndefined();
		expect(setCellBorders(view, 'right')).toBe(true);
		rows = cells(view, model);
		expect(rows[0]![1]!.borders?.right?.style).toBe('none');
		expect(rows[0]![2]!.borders?.left?.style).toBe('none');
	});

	it('draws outside and inside edges of a selection, and None clears them', () => {
		const model = sample();
		const view = viewAt(model, 'a', 'e');
		setCellBorders(view, 'outside');
		let rows = cells(view, model);
		expect(rows[0]![0]!.borders).toMatchObject({
			top: { style: 'single' },
			left: { style: 'single' },
		});
		expect(rows[0]![0]!.borders?.right).toBeUndefined();
		setCellBorders(view, 'insideV');
		rows = cells(view, model);
		expect(rows[0]![0]!.borders?.right?.style).toBe('single');
		expect(rows[0]![1]!.borders?.left?.style).toBe('single');
		setCellBorders(view, 'none');
		rows = cells(view, model);
		expect(rows[0]![0]!.borders?.top?.style).toBe('none');
		expect(rows[1]![1]!.borders?.bottom?.style).toBe('none');
	});

	it('applies dialog settings to the whole table and reads them back', () => {
		const model = sample();
		const view = viewAt(model, 'a');
		const pen = { style: 'double', sizeEighthPoints: 12, color: '#336699' };
		applyCellBorderSettings(view, {
			scope: 'table',
			sides: { top: true, bottom: true, left: true, right: true, insideH: true },
			pen,
		});
		const rows = cells(view, model);
		expect(rows[1]![2]!.borders?.bottom).toMatchObject({ style: 'double', color: '#336699' });
		expect(rows[0]![1]!.borders?.bottom?.style).toBe('double');
		expect(rows[0]![1]!.borders?.left?.style).not.toBe('double');
		const read = readCellBorderSettings(view.state, 'table')!;
		expect(read.sides).toMatchObject({ top: true, insideH: true, insideV: false });
		expect(read.pen).toMatchObject({ style: 'double', sizeEighthPoints: 12 });
	});

	it('does nothing outside a table or when read-only', () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'x' }] }];
		const doc = modelToDoc(model);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, selection: TextSelection.create(doc, 1) }),
		});
		expect(setCellBorders(view, 'all')).toBe(false);
		const tableView = viewAt(sample(), 'a');
		tableView.setProps({ editable: () => false });
		expect(setCellBorders(tableView, 'all')).toBe(false);
	});
});
