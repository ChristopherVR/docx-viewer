// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { insertPageBreak, insertColumnBreak } from './page-break-command';
import { schema } from './schema';

describe('page/column break adapter', () => {
	it('renders a page-break run as a visible marker node and round-trips it', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: 'Before' }, { text: '', break: 'page' }, { text: 'After' }],
			},
		];
		const doc = modelToDoc(model);
		const paragraph = doc.firstChild!;
		expect(paragraph.content.content.map((node) => node.type.name)).toEqual([
			'text',
			'pageBreak',
			'text',
		]);
		expect(paragraph.child(1).attrs.kind).toBe('page');
		const roundtripped = docToModel(doc, model);
		expect(roundtripped.blocks[0]).toMatchObject({
			runs: [{ text: 'Before' }, { text: '', break: 'page' }, { text: 'After' }],
		});
	});

	it('inserts a page break command at the caret without splitting the paragraph', () => {
		const paragraph = schema.nodes.paragraph.create({ id: 'p1' }, [schema.text('ab')]);
		const doc = schema.nodes.doc.create(null, [paragraph]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 2) });
		let changed = state;
		expect(insertPageBreak(state, (transaction) => (changed = state.apply(transaction)))).toBe(
			true,
		);
		expect(changed.doc.firstChild?.childCount).toBe(3);
		expect(changed.doc.firstChild?.child(1).type.name).toBe('pageBreak');
		expect(changed.doc.firstChild?.child(1).attrs.kind).toBe('page');
	});

	it('inserts a column break command with the column kind', () => {
		const paragraph = schema.nodes.paragraph.create({ id: 'p1' }, [schema.text('ab')]);
		const doc = schema.nodes.doc.create(null, [paragraph]);
		const state = EditorState.create({ doc, selection: TextSelection.create(doc, 1) });
		let changed = state;
		expect(insertColumnBreak(state, (transaction) => (changed = state.apply(transaction)))).toBe(
			true,
		);
		expect(changed.doc.firstChild?.child(0).type.name).toBe('pageBreak');
		expect(changed.doc.firstChild?.child(0).attrs.kind).toBe('column');
	});

	it('round-trips an inserted page break through DOCX save/reload', async () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'A' }] }];
		const doc = modelToDoc(model);
		const state = EditorState.create({ doc });
		const transaction = state.tr.insert(2, schema.nodes.pageBreak.create({ kind: 'page' }));
		const nextModel = docToModel(transaction.doc, model);
		const reopened = await loadDocx(await saveDocx(nextModel));
		expect(reopened.model.blocks[0]).toMatchObject({
			runs: [{ text: 'A' }, { text: '', break: 'page' }],
		});
	});
});

describe('footnote/endnote reference adapter', () => {
	it('assigns sequential note numbers by first-reference order across paragraphs', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: 'One' }, { text: '', noteReference: { kind: 'footnote', id: '7' } }],
			},
			{
				type: 'paragraph',
				id: 'p2',
				runs: [{ text: 'Two' }, { text: '', noteReference: { kind: 'footnote', id: '9' } }],
			},
		];
		const doc = modelToDoc(model);
		expect(doc.child(0).lastChild?.attrs).toMatchObject({ kind: 'footnote', id: '7', number: 1 });
		expect(doc.child(1).lastChild?.attrs).toMatchObject({ kind: 'footnote', id: '9', number: 2 });
		const roundtripped = docToModel(doc, model);
		expect(roundtripped.blocks[0]).toMatchObject({
			runs: [{ text: 'One' }, { text: '', noteReference: { kind: 'footnote', id: '7' } }],
		});
	});

	it('keeps a note-reference paragraph unchanged (same object) when an unrelated paragraph edits', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [{ text: '', noteReference: { kind: 'footnote', id: '1' } }],
			},
			{ type: 'paragraph', id: 'p2', runs: [{ text: 'Other' }] },
		];
		const doc = modelToDoc(model);
		const next = docToModel(doc, model);
		expect(next.blocks[0]).toBe(model.blocks[0]);
	});
});
