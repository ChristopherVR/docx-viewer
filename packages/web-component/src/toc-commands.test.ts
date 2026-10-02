// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph } from 'docx-core';
import { createFakeMeasurer } from '@christophervr/ooxml-core/docx/layout';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { fieldsBalanced } from './field-guard';
import { docToModel, modelToDoc } from './model-adapter';
import {
	blockPageNumbers,
	insertTableOfContents,
	removeTableOfContents,
	updateTableOfContents,
} from './toc-commands';

function sample(): DocumentModel {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'empty', runs: [] },
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Overview' }] },
		{ type: 'paragraph', id: 'b1', runs: [{ text: 'Body' }] },
		{
			type: 'paragraph',
			id: 'h2',
			style: 'Heading2',
			pageBreakBefore: true,
			runs: [{ text: 'Details' }],
		},
	];
	return model;
}

function viewFor(model: DocumentModel) {
	const doc = modelToDoc(model);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, selection: TextSelection.create(doc, 1) }),
	});
}

const entryText = (model: DocumentModel) =>
	model.blocks
		.slice(0, 2)
		.map((block) => (block as Paragraph).runs.map((run) => run.text).join(''));

describe('table of contents commands', () => {
	it('numbers blocks by the page they start on', () => {
		const numbers = blockPageNumbers(sample(), createFakeMeasurer());
		expect(numbers.get('h1')).toBe('1');
		expect(numbers.get('h2')).toBe('2');
	});

	it('inserts a TOC in place of an empty paragraph, with page numbers', () => {
		const model = sample();
		const view = viewFor(model);
		insertTableOfContents(view, model, createFakeMeasurer());
		const next = docToModel(view.state.doc, model);
		expect(entryText(next)).toEqual(['Overview\t1', 'Details\t2']);
		expect(next.blocks).toHaveLength(5);
		expect(fieldsBalanced(view.state.doc)).toBe(true);
		// Entries link to `_Toc` bookmarks added to their headings.
		const heading = next.blocks.find((block) => block.id === 'h2') as Paragraph;
		const name = heading.bookmarks?.find((bookmark) => bookmark.startsWith('_Toc'));
		expect(name).toBeTruthy();
		expect((next.blocks[1] as Paragraph).runs.find((run) => run.text)?.link).toEqual({
			anchor: name,
		});
		expect((next.blocks[0] as Paragraph).tabStops?.[0]).toMatchObject({
			align: 'right',
			leader: 'dot',
		});
	});

	it('updates the TOC after headings change and reports when there is none', () => {
		const model = sample();
		const view = viewFor(model);
		expect(updateTableOfContents(view, model, createFakeMeasurer())).toBe(false);
		insertTableOfContents(view, model, createFakeMeasurer());
		let current = docToModel(view.state.doc, model);
		const heading = current.blocks.findIndex((block) => block.id === 'h2');
		current = {
			...current,
			blocks: current.blocks.map((block, index) =>
				index === heading ? { ...(block as Paragraph), runs: [{ text: 'Specifics' }] } : block,
			),
		};
		view.updateState(EditorState.create({ doc: modelToDoc(current) }));
		expect(updateTableOfContents(view, current, createFakeMeasurer())).toBe(true);
		expect(entryText(docToModel(view.state.doc, current))).toEqual(['Overview\t1', 'Specifics\t2']);
	});

	it('limits the entries to the requested heading depth', () => {
		const model = sample();
		const view = viewFor(model);
		insertTableOfContents(view, model, createFakeMeasurer(), 1);
		const next = docToModel(view.state.doc, model);
		expect(entryText(next).slice(0, 1)).toEqual(['Overview	1']);
		expect(next.blocks.map((block) => block.id)).not.toContain('missing');
		expect((next.blocks[0] as Paragraph).runs.some((run) => run.fieldCode?.includes('"1-1"'))).toBe(
			true,
		);
		expect(next.blocks.filter((block) => (block as Paragraph).style === 'TOC2')).toHaveLength(0);
	});

	it('removes the table of contents and keeps the other blocks', () => {
		const model = sample();
		const view = viewFor(model);
		insertTableOfContents(view, model, createFakeMeasurer());
		const withToc = docToModel(view.state.doc, model);
		expect(removeTableOfContents(view, withToc)).toBe(true);
		const next = docToModel(view.state.doc, withToc);
		expect(next.blocks.map((block) => block.id)).toEqual(['h1', 'b1', 'h2']);
		expect(fieldsBalanced(view.state.doc)).toBe(true);
	});

	it('reports there is nothing to remove, and refuses a read-only view', () => {
		const model = sample();
		expect(removeTableOfContents(viewFor(model), model)).toBe(false);
		const view = viewFor(model);
		insertTableOfContents(view, model, createFakeMeasurer());
		const withToc = docToModel(view.state.doc, model);
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(removeTableOfContents(locked, withToc)).toBe(false);
	});
});
