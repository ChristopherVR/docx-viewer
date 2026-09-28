// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph, type TextRun } from '@christophervr/docx-core';
import { Slice, Fragment } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';
import { trackChangesPlugin } from './track-changes-mode';
import { at } from './test-support';

function start(text: string, second = 'Second') {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'a', runs: [{ text }] },
		{ type: 'paragraph', id: 'b', runs: [{ text: second }] },
	];
	const state = EditorState.create({
		doc: modelToDoc(model),
		plugins: [
			trackChangesPlugin(
				() => 'Ada',
				() => true,
			),
		],
	});
	return { model, state };
}

const revisions = (state: EditorState, model: ReturnType<typeof createDocument>) =>
	docToModel(state.doc, model)
		.blocks.flatMap((block) => (block as Paragraph).runs)
		.filter((run): run is TextRun & { revision: NonNullable<TextRun['revision']> } =>
			Boolean(run.revision),
		)
		.map((run) => ({ text: run.text, kind: run.revision.kind, move: run.revision.move?.name }));

const text = (value: string) => new Slice(Fragment.from(schema.text(value)), 0, 0);

describe('recording moves under Track Changes', () => {
	it('records cut then paste of the same text as a move', () => {
		let { model, state } = start('Move me here');
		// Cut "me " (positions 6–9), then paste it at the start of the second paragraph.
		state = state.apply(state.tr.delete(6, 9).setMeta('uiEvent', 'cut'));
		const target = state.doc.child(0).nodeSize + 1;
		state = state.apply(
			state.tr
				.setSelection(TextSelection.create(state.doc, target))
				.replaceSelection(text('me '))
				.setMeta('uiEvent', 'paste'),
		);
		const moves = revisions(state, model);
		expect(moves.map(({ text, kind }) => [text, kind])).toEqual([
			['me ', 'moveFrom'],
			['me ', 'moveTo'],
		]);
		const [firstMove, secondMove] = [at(moves, 0), at(moves, 1)];
		expect(firstMove.move).toMatch(/^move\d+$/);
		expect(secondMove.move).toBe(firstMove.move);
	});

	it('records a drag-and-drop move in one transaction', () => {
		let { model, state } = start('Drag this word', 'Target');
		const target = state.doc.child(0).nodeSize + 1 + 'Target'.length;
		// A drop deletes the source, then inserts at the target (positions after the deletion).
		const tr = state.tr.delete(6, 11);
		tr.insert(tr.mapping.map(target), schema.text('this '));
		state = state.apply(tr.setMeta('uiEvent', 'drop'));
		const moves = revisions(state, model);
		expect(moves.map(({ text, kind }) => [text, kind])).toEqual([
			['this ', 'moveFrom'],
			['this ', 'moveTo'],
		]);
		expect(at(moves, 0).move).toBe(at(moves, 1).move);
		expect(state.doc.child(1).textContent).toBe('Targetthis ');
	});

	it('keeps a paste of different text a plain insertion', () => {
		let { model, state } = start('Keep me');
		state = state.apply(state.tr.delete(6, 8).setMeta('uiEvent', 'cut'));
		state = state.apply(state.tr.insert(1, schema.text('Other ')).setMeta('uiEvent', 'paste'));
		expect(revisions(state, model).map(({ kind }) => kind)).toEqual(['insert', 'delete']);
	});
});
