// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type DocumentModel, type Paragraph } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { docToModel, modelToDoc } from './model-adapter';
import { acceptRevisionRange, collectRevisionRanges, rejectRevisionRange } from './review-commands';
import { at, paragraphAt } from './test-support';

function movedModel(): DocumentModel {
	const model = createDocument();
	const side = { author: 'Ada', move: { name: 'move1', rangeId: '10' } };
	model.blocks = [
		{
			type: 'paragraph',
			id: 'a',
			runs: [
				{ text: 'Keep ' },
				{ text: 'moved', revision: { ...side, id: '1', kind: 'moveFrom' } },
			],
		},
		{
			type: 'paragraph',
			id: 'b',
			runs: [
				{
					text: 'moved',
					revision: { ...side, id: '2', kind: 'moveTo', move: { name: 'move1', rangeId: '11' } },
				},
				{ text: ' here' },
				{ text: ' new', revision: { author: 'Ada', id: '3', kind: 'insert' } },
			],
		},
	];
	return model;
}

const view = (model: DocumentModel) =>
	new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc(model) }),
	});
const texts = (model: DocumentModel) =>
	model.blocks.map((block) => (block as Paragraph).runs.map((run) => run.text).join(''));

describe('tracked moves in the editor', () => {
	it('keeps move kinds and names through the editor', () => {
		const model = movedModel();
		const editor = view(model);
		expect(collectRevisionRanges(editor.state.doc).map((range) => range.move)).toEqual([
			'move1',
			'move1',
			undefined,
		]);
		expect(editor.dom.querySelectorAll('.dve-revision-move')).toHaveLength(2);
		const round = docToModel(editor.state.doc, model);
		expect(at(paragraphAt(round.blocks, 1).runs, 0).revision).toEqual(
			at(paragraphAt(model.blocks, 1).runs, 0).revision,
		);
	});

	it('accepts or rejects both sides of a move together, leaving other changes', () => {
		const model = movedModel();
		const accepting = view(model);
		acceptRevisionRange(accepting, at(collectRevisionRanges(accepting.state.doc), 1));
		const accepted = docToModel(accepting.state.doc, model);
		expect(texts(accepted)).toEqual(['Keep ', 'moved here new']);
		expect(collectRevisionRanges(accepting.state.doc).map((range) => range.id)).toEqual(['3']);

		const rejecting = view(model);
		rejectRevisionRange(rejecting, at(collectRevisionRanges(rejecting.state.doc), 0));
		expect(texts(docToModel(rejecting.state.doc, model))).toEqual(['Keep moved', ' here new']);
	});
});
