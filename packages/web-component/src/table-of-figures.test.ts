// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph } from '@christophervr/docx-core';
import { createFakeMeasurer } from '@christophervr/ooxml-core/docx/layout';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import {
	insertTableOfContents,
	removeTableOfContents,
	updateTableOfContents,
} from './toc-commands';

const caption = (id: string, label: string, number: string, text: string): Paragraph => ({
	type: 'paragraph',
	id,
	runs: [
		{ text: `${label} ` },
		{ text: number, field: { instr: ` SEQ ${label} \* ARABIC `, simple: true } },
		{ text: `: ${text}` },
	],
});

function sample(): DocumentModel {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'empty', runs: [] },
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Overview' }] },
		caption('c1', 'Figure', '1', 'Chart'),
		caption('t1', 'Table', '1', 'Totals'),
		{ ...caption('c2', 'Figure', '2', 'Map'), pageBreakBefore: true },
	];
	return model;
}
function viewFor(model: DocumentModel) {
	const doc = modelToDoc(model);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1) }),
	});
}
const texts = (model: DocumentModel) =>
	model.blocks.map((block) => (block as Paragraph).runs.map((run) => run.text).join(''));

describe('table of figures', () => {
	it('lists only the captions of its label, with pages, and does not disturb a heading TOC', () => {
		const model = sample();
		const view = viewFor(model);
		insertTableOfContents(view, model, createFakeMeasurer(), 3, 'Figure');
		const next = docToModel(view.state.doc, model);
		expect(texts(next).slice(0, 2)).toEqual(['Figure 1: Chart\t1', 'Figure 2: Map\t2']);
		const first = next.blocks[0] as Paragraph;
		expect(first.runs.some((run) => run.fieldCode?.includes('\c "Figure"'))).toBe(true);
		// A heading TOC lookup ignores it; removing figures leaves the captions (the empty paragraph it replaced is gone).
		expect(removeTableOfContents(view, next)).toBe(false);
		expect(removeTableOfContents(view, next, 'Figure')).toBe(true);
		expect(texts(docToModel(view.state.doc, next))).toEqual(texts(model).slice(1));
	});

	it('updates after a caption is added and reports when there is none', () => {
		const model = sample();
		const view = viewFor(model);
		expect(updateTableOfContents(view, model, createFakeMeasurer(), 'Figure')).toBe(false);
		insertTableOfContents(view, model, createFakeMeasurer(), 3, 'Table');
		let current = docToModel(view.state.doc, model);
		expect(texts(current)[0]).toBe('Table 1: Totals\t1');
		current = { ...current, blocks: [...current.blocks, caption('t2', 'Table', '2', 'More')] };
		view.updateState(EditorState.create({ doc: modelToDoc(current) }));
		expect(updateTableOfContents(view, current, createFakeMeasurer(), 'Table')).toBe(true);
		expect(texts(docToModel(view.state.doc, current)).slice(0, 2)).toEqual([
			'Table 1: Totals\t1',
			'Table 2: More\t2',
		]);
	});
});
