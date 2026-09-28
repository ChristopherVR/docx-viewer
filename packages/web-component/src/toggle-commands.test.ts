// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { runStylesPlugin } from './run-styles';
import { toggleFormat } from './toggle-commands';
import { at, paragraphAt } from './test-support';

function setup(runs: Paragraph['runs'], style?: string) {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs, ...(style ? { style } : {}) }];
	const doc = modelToDoc(model);
	const state = EditorState.create({
		doc,
		plugins: [runStylesPlugin(() => model)],
		selection: TextSelection.create(doc, 1, doc.firstChild!.nodeSize - 1),
	});
	return { model, state };
}
const apply = (state: EditorState, key: 'bold' | 'italic') => {
	let next = state;
	toggleFormat(key)(state, (tr) => (next = state.apply(tr)));
	return next;
};
const runOf = (state: EditorState, model: ReturnType<typeof createDocument>) =>
	at(paragraphAt(docToModel(state.doc, model).blocks, 0).runs, 0);

describe('style-aware Bold', () => {
	it('un-bolds heading text with an explicit off, and bolds it again', () => {
		// Heading 1 in the default styles is not bold; Title is not either, so use a bold char style.
		const { model, state } = setup([{ text: 'Heading text', style: 'Strong' }]);
		model.characterStyles!.styles.Strong = {
			id: 'Strong',
			type: 'character',
			formatting: { bold: true },
		};
		const off = apply(state, 'bold');
		expect(runOf(off, model)).toMatchObject({ text: 'Heading text', bold: false });
		const on = apply(off, 'bold');
		expect(runOf(on, model).bold).toBe(true);
	});

	it('toggles plain text with a mark and removes it again without an explicit off', () => {
		const { model, state } = setup([{ text: 'Plain' }]);
		const bolded = apply(state, 'bold');
		expect(runOf(bolded, model).bold).toBe(true);
		expect(runOf(apply(bolded, 'bold'), model).bold).toBeUndefined();
	});
});
