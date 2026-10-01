// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { insertCaption, seqInstruction } from './caption-commands';
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

describe('inserting a caption', () => {
	it('renumbers captions and refreshes REF results that quote them, but not other references', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const doc = schema.node('doc', null, [
			paragraph('intro', [schema.text('Intro')]),
			paragraph(
				'cap',
				[schema.text('Figure '), field('1', seqInstruction('Figure')), schema.text(': Chart')],
				['_RefCap'],
			),
			paragraph('head', [schema.text('Heading text')], ['_RefHead']),
			paragraph('refs', [
				field('Figure 1', ' REF _RefCap \\h '),
				schema.text('|'),
				field('stale heading', ' REF _RefHead \\h '),
			]),
		]);
		const view = new EditorView(host, {
			state: EditorState.create({ doc, schema, selection: TextSelection.create(doc, 2) }),
		});
		// A new figure above the existing one pushes it to Figure 2.
		insertCaption(view, { label: 'Figure', labelText: 'Figure', text: 'New', position: 'below' });
		const texts: string[] = [];
		view.state.doc.forEach((node) => texts.push(node.textContent));
		expect(texts[1]).toBe('Figure 1: New');
		expect(texts[2]).toBe('Figure 2: Chart');
		expect(texts[4]).toBe('Figure 2|stale heading');
	});
});
