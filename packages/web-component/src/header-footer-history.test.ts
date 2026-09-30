// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { closeHistory, history, undo, redo } from 'prosemirror-history';
import { describe, expect, it } from 'vitest';
import { sectionPartsJson, SectionPartsStep } from './header-footer-history';
import { modelToDoc, docToModel } from './model-adapter';
import { withPageNumber } from './header-footer-commands';

describe('header/footer content history', () => {
	it('groups consecutive content snapshots and keeps body positions unchanged', () => {
		let state = EditorState.create({ doc: modelToDoc(createDocument()), plugins: [history()] });
		const initial = state.doc.attrs.sectionParts;
		const first = new SectionPartsStep('[{"endsAtBlockId":"a"}]');
		expect(first.getMap().map(12)).toBe(12);
		state = state.apply(closeHistory(state.tr).step(first));
		state = state.apply(state.tr.step(new SectionPartsStep('[{"endsAtBlockId":"b"}]')));
		undo(state, (tr) => {
			state = state.apply(tr);
		});
		expect(state.doc.attrs.sectionParts).toBe(initial);
		redo(state, (tr) => {
			state = state.apply(tr);
		});
		expect(state.doc.attrs.sectionParts).toContain('"b"');
	});
	it('records creation and content, retains body edits and restores content on undo/redo', () => {
		let model = createDocument();
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc: modelToDoc(model), plugins: [history()] }),
			dispatchTransaction(tr) {
				view.updateState(view.state.apply(tr));
				model = docToModel(view.state.doc, model);
			},
		});
		const inserted = withPageNumber(model, 'bottom', 'center', () => 'footer-p');
		view.dispatch(
			closeHistory(view.state.tr)
				.setDocAttribute('sectionParts', sectionPartsJson(inserted.sections))
				.setDocAttribute('sections', JSON.stringify(inserted.sections)),
		);
		expect(model.sections![0]!.footers!.default!.blocks[0]!.id).toBe('footer-p');
		view.dispatch(closeHistory(view.state.tr).insertText('Body', 1));
		undo(view.state, view.dispatch);
		expect(
			model.blocks[0]!.type === 'paragraph' &&
				model.blocks[0]!.runs.map((run) => run.text).join(''),
		).toBe('');
		expect(model.sections![0]!.footers).toBeDefined();
		undo(view.state, view.dispatch);
		expect(model.sections![0]!.footers).toBeUndefined();
		redo(view.state, view.dispatch);
		expect(model.sections![0]!.footers!.default!.partName).toBe('word/footer1.xml');
		view.destroy();
	});
});
