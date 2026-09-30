// @vitest-environment jsdom
import { createDocument, type DocumentModel } from '@christophervr/docx-core';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { modelToDoc, docToModel } from './model-adapter';
import { runStylesPlugin, runFormattingCss } from './run-styles';
import { toggleVerticalAlign } from './inline-commands';
import { applyFontFormat } from './font-format';
import { selectionScript } from './script-state';
import { syncFormatControls } from './ribbon-controls';

function setup(model: DocumentModel) {
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [runStylesPlugin(() => model), history()],
		}),
	});
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);
	return view;
}
const scriptModel = () => {
	const model = createDocument();
	model.characterStyles = {
		docDefaults: { verticalAlign: 'superscript', fontSize: 12 },
		styles: {},
		warnings: [],
	};
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Inherited' }] }];
	return model;
};
describe('script formatting reset', () => {
	it('resolves paragraph style ancestry and character overrides before toggling', () => {
		const model = createDocument();
		model.paragraphStyles = {
			docDefaults: {},
			styles: {
				Parent: { id: 'Parent', name: 'Parent', formatting: {} },
				Child: { id: 'Child', name: 'Child', basedOn: 'Parent', formatting: {} },
			},
			warnings: [],
		};
		model.characterStyles = {
			docDefaults: {},
			styles: {
				Parent: { id: 'Parent', type: 'paragraph', formatting: { verticalAlign: 'superscript' } },
				Ref: { id: 'Ref', type: 'character', formatting: { verticalAlign: 'subscript' } },
			},
			warnings: [],
		};
		model.blocks = [
			{ type: 'paragraph', id: 'p', style: 'Child', runs: [{ text: 'Reference', style: 'Ref' }] },
		];
		const view = setup(model);
		expect(selectionScript(view.state)).toBe('subscript');
		toggleVerticalAlign(view, 'subscript');
		expect(selectionScript(view.state)).toBe('none');
		const p = docToModel(view.state.doc, model).blocks[0]!;
		if (p.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(p.runs[0]).toMatchObject({ style: 'Ref', verticalAlign: 'baseline' });
		view.destroy();
	});
	it('shows inherited script on the ribbon and resets it with one undoable baseline override', () => {
		const model = scriptModel();
		const view = setup(model);
		const toolbar = document.createElement('div');
		toolbar.innerHTML =
			'<button aria-label="Superscript"></button><button aria-label="Subscript"></button>';
		syncFormatControls(toolbar, view.state);
		expect(toolbar.firstElementChild!.getAttribute('aria-pressed')).toBe('true');
		toggleVerticalAlign(view, 'superscript');
		expect(selectionScript(view.state)).toBe('none');
		const p = docToModel(view.state.doc, model).blocks[0]!;
		if (p.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(p.runs[0]!.verticalAlign).toBe('baseline');
		expect(runFormattingCss({ fontSize: 12, verticalAlign: 'baseline' })).toContain(
			'font-size:12pt',
		);
		undo(view.state, view.dispatch);
		expect(selectionScript(view.state)).toBe('superscript');
		view.destroy();
	});
	it('normalizes a mixed inherited/explicit selection through the Font dialog command', () => {
		const model = scriptModel();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				runs: [
					{ text: 'Inherited' },
					{ text: 'Plain', verticalAlign: 'baseline' },
					{ text: 'Sub', verticalAlign: 'subscript' },
				],
			},
		];
		const view = setup(model);
		expect(selectionScript(view.state)).toBeNull();
		applyFontFormat(view, { script: 'none' });
		expect(selectionScript(view.state)).toBe('none');
		undo(view.state, view.dispatch);
		expect(selectionScript(view.state)).toBeNull();
		view.destroy();
	});
	it('resets inherited script on a collapsed caret while retaining inheritance on existing text', () => {
		const model = scriptModel();
		const view = setup(model);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1)));
		toggleVerticalAlign(view, 'superscript');
		expect(selectionScript(view.state)).toBe('none');
		view.dispatch(view.state.tr.insertText('New'));
		const p = docToModel(view.state.doc, model).blocks[0]!;
		if (p.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(p.runs).toEqual([{ text: 'New', verticalAlign: 'baseline' }, { text: 'Inherited' }]);
		view.destroy();
	});
});
