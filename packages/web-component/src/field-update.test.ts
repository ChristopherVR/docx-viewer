// @vitest-environment jsdom
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { undo } from 'prosemirror-history';
import { history } from 'prosemirror-history';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { seqInstruction } from './caption-commands';
import { updateFields } from './field-update';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => (document.body.innerHTML = ''));

const field = (text: string, instr: string) =>
	schema.text(text, [schema.marks.field!.create({ instr, simple: true })]);
const paragraph = (id: string, content: unknown[], bookmarks: string[] = []) =>
	schema.nodes.paragraph!.create({ id, bookmarks }, content as never);

function editor(...blocks: ReturnType<typeof paragraph>[]) {
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, {
		state: EditorState.create({
			doc: schema.node('doc', null, blocks),
			schema,
			plugins: [history()],
		}),
	});
}
const texts = (view: EditorView) => {
	const out: string[] = [];
	view.state.doc.forEach((node) => out.push(node.textContent));
	return out;
};
const pages = (id: string) => ({ p0: '1', p1: '4' })[id];

describe('updateFields', () => {
	it('renumbers captions, then refreshes REF and PAGEREF results', () => {
		const view = editor(
			paragraph('p0', [schema.text('Intro ')], []),
			paragraph(
				'p1',
				[schema.text('Figure '), field('7', seqInstruction('Figure')), schema.text(': Chart')],
				['_Ref1'],
			),
			paragraph('p2', [
				schema.text('See '),
				field('Figure 9', ' REF _Ref1 \h '),
				schema.text(' on page '),
				field('9', ' PAGEREF _Ref1 \h '),
			]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)).toEqual(['Intro ', 'Figure 1: Chart', 'See Figure 1 on page 4']);
	});

	it('follows a heading text REF, keeps unsupported switches and missing targets, and undoes in one step', () => {
		const view = editor(
			paragraph('p0', [schema.text('New title')], ['_Ref2']),
			paragraph('p1', [
				field('Old title', ' REF _Ref2 \h '),
				schema.text('|'),
				field('x', ' REF _Ref2 \n \h '),
				schema.text('|'),
				field('gone', ' REF _Missing \h '),
			]),
		);
		expect(updateFields(view, pages)).toBe(true);
		expect(texts(view)[1]).toBe('New title|x|gone');
		expect(updateFields(view, pages)).toBe(false);
		undo(view.state, view.dispatch);
		expect(texts(view)[1]).toBe('Old title|x|gone');
	});

	it('does nothing in a read-only view', () => {
		const view = editor(paragraph('p0', [field('1', seqInstruction('Figure'))]));
		view.setProps({ editable: () => false });
		expect(updateFields(view, pages)).toBe(false);
	});
});
