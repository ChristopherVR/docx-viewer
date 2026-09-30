import { describe, expect, it } from 'vitest';
import { createDocument, halfPoints } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { schema } from './schema';

describe('advanced run editor preservation', () => {
	it('keeps scale, kerning and baseline position when a text edit splits an imported run', () => {
		const model = createDocument();
		const formatting = {
			textScalePercent: 125,
			kerningHalfPoints: halfPoints(24),
			positionHalfPoints: halfPoints(-6),
		};
		model.blocks = [
			{ type: 'paragraph', id: 'advanced', runs: [{ text: 'Advanced', ...formatting }] },
		];
		const state = EditorState.create({ schema, doc: modelToDoc(model) });
		const edited = state.apply(state.tr.insertText(' changed', 9));
		const result = docToModel(edited.doc, model);
		const paragraph = result.blocks[0];
		expect(paragraph?.type).toBe('paragraph');
		if (paragraph?.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs.map((run) => run.text).join('')).toBe('Advanced changed');
		for (const run of paragraph.runs) expect(run).toMatchObject(formatting);
	});
});
