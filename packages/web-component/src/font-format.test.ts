// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { applyFontFormat, readFontFormat } from './font-format';
import { runStylesPlugin } from './run-styles';
import { schema } from './schema';

function editor(
	...runs: Array<{ text: string; marks?: ReturnType<typeof schema.marks.bold.create>[] }>
) {
	const model = createDocument();
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create(
			{ id: 'p' },
			runs.map((run) => schema.text(run.text, run.marks ?? [])),
		),
	]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema, plugins: [runStylesPlugin(() => model)] }),
	});
	return view;
}
const select = (view: EditorView, from: number, to: number) =>
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
const selectAll = (view: EditorView) => select(view, 1, view.state.doc.content.size - 1);
const extras = (view: EditorView, pos = 2) =>
	(schema.marks.runProperties!.isInSet(view.state.doc.nodeAt(pos)!.marks)?.attrs.props ??
		{}) as Record<string, unknown>;

describe('readFontFormat', () => {
	it('reports defaults for plain text', () => {
		const view = editor({ text: 'plain' });
		selectAll(view);
		expect(readFontFormat(view.state)).toMatchObject({
			family: 'Calibri',
			size: 11,
			bold: false,
			underline: 'none',
			script: 'none',
			spacing: 0,
			hidden: false,
		});
	});

	it('reads direct marks and extra run properties', () => {
		const view = editor({
			text: 'fancy',
			marks: [
				schema.marks.bold!.create(),
				schema.marks.underline!.create(),
				schema.marks.font!.create({ family: 'Georgia', size: 14, color: '#c00000' }),
				schema.marks.runProperties!.create({
					props: { underlineStyle: 'wave', smallCaps: true, characterSpacingTwips: 40 },
				}),
			],
		});
		selectAll(view);
		expect(readFontFormat(view.state)).toMatchObject({
			family: 'Georgia',
			size: 14,
			color: '#c00000',
			bold: true,
			underline: 'wave',
			smallCaps: true,
			spacing: 2,
		});
	});

	it('marks fields that differ across the selection as mixed (null)', () => {
		const view = editor({ text: 'one ', marks: [schema.marks.bold!.create()] }, { text: 'two' });
		selectAll(view);
		const format = readFontFormat(view.state);
		expect(format.bold).toBeNull();
		expect(format.family).toBe('Calibri');
	});
});

describe('applyFontFormat', () => {
	it('applies family, size, colour and toggles, leaving other text alone', () => {
		const view = editor({ text: 'abcdef' });
		select(view, 1, 4);
		applyFontFormat(view, {
			family: 'Verdana',
			size: 16,
			color: '#0070c0',
			bold: true,
			italic: true,
		});
		const first = view.state.doc.nodeAt(2)!;
		expect(first.marks.map((mark) => mark.type.name)).toEqual(
			expect.arrayContaining(['bold', 'italic', 'font']),
		);
		expect(first.marks.find((mark) => mark.type.name === 'font')!.attrs).toMatchObject({
			family: 'Verdana',
			size: 16,
			color: '#0070c0',
		});
		expect(view.state.doc.nodeAt(5)!.marks).toHaveLength(0);
	});

	it('turns a toggle off across a mixed selection', () => {
		const view = editor({ text: 'one ', marks: [schema.marks.bold!.create()] }, { text: 'two' });
		selectAll(view);
		applyFontFormat(view, { bold: false });
		expect(readFontFormat(view.state).bold).toBe(false);
	});

	it('sets underline kinds, keeping the underline mark and recording the style', () => {
		const view = editor({ text: 'under' });
		selectAll(view);
		applyFontFormat(view, { underline: 'double', underlineColor: '#ff0000' });
		expect(readFontFormat(view.state).underline).toBe('double');
		expect(extras(view)).toMatchObject({ underlineStyle: 'double', underlineColor: '#ff0000' });
		applyFontFormat(view, { underline: 'single', underlineColor: null });
		expect(readFontFormat(view.state).underline).toBe('single');
		expect(extras(view)).toEqual({});
		applyFontFormat(view, { underline: 'none' });
		expect(readFontFormat(view.state).underline).toBe('none');
	});

	it('records effects and character spacing, and removes them again', () => {
		const view = editor({ text: 'effects' });
		selectAll(view);
		applyFontFormat(view, { smallCaps: true, hidden: true, doubleStrike: true, spacing: -1.5 });
		expect(extras(view)).toEqual({
			smallCaps: true,
			vanish: true,
			doubleStrike: true,
			characterSpacingTwips: -30,
		});
		applyFontFormat(view, { smallCaps: false, hidden: false, doubleStrike: false, spacing: 0 });
		expect(extras(view)).toEqual({
			characterSpacingTwips: 0,
			smallCaps: false,
			vanish: false,
			doubleStrike: false,
		});
	});
	it('turns inherited font effects off with explicit overrides', () => {
		const model = createDocument();
		model.characterStyles = {
			docDefaults: { caps: true, smallCaps: true, doubleStrike: true, vanish: true },
			styles: {},
			warnings: [],
		};
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({
				schema,
				doc: schema.node('doc', null, [
					schema.nodes.paragraph!.create({ id: 'p' }, schema.text('Inherited')),
				]),
				plugins: [runStylesPlugin(() => model)],
			}),
		});
		selectAll(view);
		expect(readFontFormat(view.state)).toMatchObject({
			caps: true,
			smallCaps: true,
			doubleStrike: true,
			hidden: true,
		});
		applyFontFormat(view, { caps: false, smallCaps: false, doubleStrike: false, hidden: false });
		expect(readFontFormat(view.state)).toMatchObject({
			caps: false,
			smallCaps: false,
			doubleStrike: false,
			hidden: false,
		});
		expect(extras(view)).toEqual({
			caps: false,
			smallCaps: false,
			doubleStrike: false,
			vanish: false,
		});
		view.destroy();
	});

	it('switches between superscript and subscript', () => {
		const view = editor({ text: 'script' });
		selectAll(view);
		applyFontFormat(view, { script: 'superscript' });
		expect(readFontFormat(view.state).script).toBe('superscript');
		applyFontFormat(view, { script: 'subscript' });
		expect(readFontFormat(view.state).script).toBe('subscript');
		applyFontFormat(view, { script: 'none' });
		expect(readFontFormat(view.state).script).toBe('none');
	});

	it('changes the stored marks for a collapsed caret', () => {
		const view = editor({ text: 'caret' });
		select(view, 3, 3);
		applyFontFormat(view, { bold: true, caps: true });
		const stored = view.state.storedMarks!;
		expect(stored.some((mark) => mark.type.name === 'bold')).toBe(true);
		expect(stored.find((mark) => mark.type.name === 'runProperties')!.attrs.props).toEqual({
			caps: true,
		});
	});

	it('does nothing in a read-only view', () => {
		const model = createDocument();
		const base = editor({ text: 'locked' });
		const view = new EditorView(document.createElement('div'), {
			state: base.state,
			editable: () => false,
		});
		void model;
		selectAll(view);
		applyFontFormat(view, { bold: true });
		expect(view.state.doc.nodeAt(2)!.marks).toHaveLength(0);
	});
});
