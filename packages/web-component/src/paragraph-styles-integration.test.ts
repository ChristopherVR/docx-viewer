// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { DocumentModel, ParagraphStyleCatalog } from '@christophervr/docx-core';
import { syncStylePicker, paragraphStylesPlugin } from './paragraph-styles';
import { schema } from './schema';
import { paragraphAt } from './test-support';

const catalog: ParagraphStyleCatalog = {
	docDefaults: {},
	styles: {
		Body: { id: 'Body', name: 'Body', isDefault: true, formatting: { align: 'left' } },
		Centered: { id: 'Centered', name: 'Centered', formatting: { align: 'center' } },
	},
	warnings: [],
};

function model(): DocumentModel {
	return {
		blocks: [{ type: 'paragraph', id: 'p1', style: 'Body', runs: [{ text: 'Text' }] }],
		paragraphStyles: catalog,
		page: {
			width: 816,
			height: 1056,
			marginTop: 96,
			marginRight: 96,
			marginBottom: 96,
			marginLeft: 96,
		},
		warnings: [],
	};
}

describe('paragraph style editor integration', () => {
	it('applies a chosen paragraph style at a collapsed text cursor', () => {
		let documentModel = model();
		const doc = schema.node('doc', null, [
			schema.node('paragraph', { id: 'p1', style: 'Body' }, [schema.text('Text')]),
		]);
		const state = EditorState.create({
			doc,
			selection: TextSelection.create(doc, 2),
			plugins: [paragraphStylesPlugin(() => documentModel)],
		});
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state,
			dispatchTransaction(transaction) {
				const next = view.state.apply(transaction);
				view.updateState(next);
				documentModel = {
					...documentModel,
					blocks: [{ ...documentModel.blocks[0], style: next.doc.firstChild?.attrs.style }],
				} as DocumentModel;
			},
		});
		const toolbar = document.createElement('div');
		toolbar.innerHTML = '<div id="dve-panel-home"></div>';
		syncStylePicker(toolbar, view, documentModel, 'en');
		const select = toolbar.querySelector<HTMLSelectElement>('[data-paragraph-styles]')!;
		select.value = 'Centered';
		select.dispatchEvent(new Event('change', { bubbles: true }));
		expect(view.state.doc.firstChild?.attrs.style).toBe('Centered');
		expect(paragraphAt(documentModel.blocks, 0).style).toBe('Centered');
		expect(view.dom.querySelector('p')?.style.textAlign).toBe('center');
		view.destroy();
		host.remove();
	});
});
