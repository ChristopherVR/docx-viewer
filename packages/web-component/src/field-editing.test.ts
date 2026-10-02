// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph, type TextRun } from 'docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import {
	deleteAroundMarkers,
	fieldGuardPlugin,
	fieldsBalanced,
	replaceAroundMarkers,
} from './field-guard';
import { docToModel, modelToDoc } from './model-adapter';

const fieldRuns: TextRun[] = [
	{ text: 'See ' },
	{ text: '', fieldChar: 'begin', bold: true },
	{ text: '', fieldCode: ' REF _Ref1 \\h ' },
	{ text: '', fieldChar: 'separate' },
	{ text: 'Figure 1', field: { instr: 'REF _Ref1 \\h' } },
	{ text: '', fieldChar: 'end' },
	{ text: ' above.' },
];

function fieldModel() {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: fieldRuns }];
	return model;
}

describe('editable fields in the editor', () => {
	it('maps field markers to hidden nodes and result text to a field mark, round-tripping both', () => {
		const model = fieldModel();
		const doc = modelToDoc(model);
		const paragraph = doc.firstChild!;
		expect(paragraph.textContent).toBe('See Figure 1 above.');
		const kinds: string[] = [];
		paragraph.forEach((node) => kinds.push(node.type.name));
		expect(kinds).toEqual([
			'text',
			'fieldMarker',
			'fieldMarker',
			'fieldMarker',
			'text',
			'fieldMarker',
			'text',
		]);
		expect(paragraph.child(4).marks.map((mark) => mark.type.name)).toContain('field');
		expect((docToModel(doc, model).blocks[0] as Paragraph).runs).toEqual(fieldRuns);
	});

	it('edits the field result without losing the field code', () => {
		const model = fieldModel();
		const doc = modelToDoc(model);
		let state = EditorState.create({ doc, plugins: [fieldGuardPlugin()] });
		// "See " is 4 characters, then three markers; the result text starts at position 1 + 4 + 3.
		const start = 1 + 4 + 3;
		state = state.apply(state.tr.insertText('Table 3', start, start + 'Figure 1'.length));
		const runs = (docToModel(state.doc, model).blocks[0] as Paragraph).runs;
		expect(runs[4]).toEqual({ text: 'Table 3', field: { instr: 'REF _Ref1 \\h' } });
		expect(runs.filter((run) => run.fieldChar).map((run) => run.fieldChar)).toEqual([
			'begin',
			'separate',
			'end',
		]);
	});

	it('rejects deleting part of a field but allows deleting a whole field', () => {
		const doc = modelToDoc(fieldModel());
		const state = EditorState.create({ doc, plugins: [fieldGuardPlugin()] });
		expect(fieldsBalanced(state.doc)).toBe(true);
		// Deleting from inside the result through " above." removes only the end marker.
		const partial = state.apply(state.tr.delete(1 + 4 + 3 + 2, state.doc.firstChild!.nodeSize - 3));
		expect(partial.doc.eq(state.doc)).toBe(true);
		// Deleting the whole field, markers included, is allowed.
		const end = 1 + 4 + 3 + 'Figure 1'.length + 1;
		const whole = state.apply(state.tr.delete(1 + 4, end));
		expect(whole.doc.firstChild!.textContent).toBe('See  above.');
		expect(fieldsBalanced(whole.doc)).toBe(true);
	});

	it('does not merge typed text after a field into its result', () => {
		const model = fieldModel();
		const doc = modelToDoc(model);
		let state = EditorState.create({ doc, plugins: [fieldGuardPlugin()] });
		const afterEnd = 1 + 4 + 3 + 'Figure 1'.length + 1;
		state = state.apply(
			state.tr.setSelection(TextSelection.create(state.doc, afterEnd)).insertText('!'),
		);
		const runs = (docToModel(state.doc, model).blocks[0] as Paragraph).runs;
		expect(runs.at(-1)).toEqual({ text: '! above.' });
	});
});

describe('field selection', () => {
	it('trims hidden markers from the edges of a selection', () => {
		const doc = modelToDoc(fieldModel());
		let state = EditorState.create({ doc, plugins: [fieldGuardPlugin()] });
		const resultStart = 1 + 4 + 3;
		const resultEnd = resultStart + 'Figure 1'.length;
		state = state.apply(
			state.tr.setSelection(TextSelection.create(state.doc, 1 + 4, resultEnd + 1)),
		);
		expect([state.selection.from, state.selection.to]).toEqual([resultStart, resultEnd]);
	});
});

describe('editing across field boundaries', () => {
	const resultStart = 1 + 4 + 3;
	const resultEnd = resultStart + 'Figure 1'.length;
	const text = (state: EditorState) => state.doc.firstChild!.textContent;

	it('types over a selection that crosses the field end, keeping the markers', () => {
		const state = EditorState.create({
			doc: modelToDoc(fieldModel()),
			plugins: [fieldGuardPlugin()],
		});
		const next = state.apply(replaceAroundMarkers(state, resultStart, resultEnd + 2, 'Map')!);
		expect(text(next)).toBe('See Mapabove.');
		expect(fieldsBalanced(next.doc)).toBe(true);
		const runs = (docToModel(next.doc, fieldModel()).blocks[0] as Paragraph).runs;
		expect(runs[4]).toMatchObject({ text: 'Map', field: { instr: 'REF _Ref1 \\h' } });
	});

	it('backspaces over the end marker into the result, and Delete over the begin markers', () => {
		const state = EditorState.create({
			doc: modelToDoc(fieldModel()),
			plugins: [fieldGuardPlugin()],
		});
		const afterEnd = state.apply(
			state.tr.setSelection(TextSelection.create(state.doc, resultEnd + 1)),
		);
		const backspaced = afterEnd.apply(deleteAroundMarkers(afterEnd, true)!);
		expect(text(backspaced)).toBe('See Figure  above.');
		const beforeBegin = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1 + 4)));
		const deleted = beforeBegin.apply(deleteAroundMarkers(beforeBegin, false)!);
		expect(text(deleted)).toBe('See igure 1 above.');
		expect(fieldsBalanced(deleted.doc)).toBe(true);
	});
});
