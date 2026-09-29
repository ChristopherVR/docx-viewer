// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { schema } from './schema';
import { sortParagraphs } from './sort-commands';

function editor(...texts: string[]) {
	const doc = schema.node(
		'doc',
		null,
		texts.map((text, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}` }, text ? schema.text(text) : undefined),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const all = (view: EditorView) =>
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);
const texts = (view: EditorView) => {
	const out: string[] = [];
	view.state.doc.forEach((node) => out.push(node.textContent));
	return out;
};

describe('sortParagraphs', () => {
	it('sorts ascending and descending, comparing numbers as numbers', () => {
		const view = editor('Item 10', 'banana', 'Item 2', 'Apple');
		all(view);
		expect(sortParagraphs(view, 'ascending', 'en')).toBe(true);
		expect(texts(view)).toEqual(['Apple', 'banana', 'Item 2', 'Item 10']);
		expect(sortParagraphs(view, 'descending', 'en')).toBe(true);
		expect(texts(view)).toEqual(['Item 10', 'Item 2', 'banana', 'Apple']);
	});

	it("keeps each paragraph's attributes and only touches the selected ones", () => {
		const view = editor('c', 'b', 'a');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 4, 7)));
		expect(sortParagraphs(view, 'ascending')).toBe(true);
		expect(texts(view)).toEqual(['c', 'a', 'b']);
		expect(view.state.doc.child(1).attrs.id).toBe('p2');
	});

	it('does nothing for one paragraph, an already sorted list, or a read-only view', () => {
		const single = editor('only');
		all(single);
		expect(sortParagraphs(single, 'ascending')).toBe(false);
		const sorted = editor('a', 'b');
		all(sorted);
		expect(sortParagraphs(sorted, 'ascending')).toBe(false);
		const locked = new EditorView(document.createElement('div'), {
			state: editor('b', 'a').state,
			editable: () => false,
		});
		expect(sortParagraphs(locked, 'ascending')).toBe(false);
	});

	it('refuses a selection that includes a table', () => {
		const cell = schema.nodes.tableCell!.create(
			{},
			schema.nodes.paragraph!.create({ id: 'c' }, schema.text('x')),
		);
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'a' }, schema.text('z')),
			schema.nodes.table!.create({ id: 't' }, schema.nodes.tableRow!.create({}, cell)),
			schema.nodes.paragraph!.create({ id: 'b' }, schema.text('a')),
		]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		all(view);
		expect(sortParagraphs(view, 'ascending')).toBe(false);
	});

	it('is reachable from the Sort menu', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const menu = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Sort"]')!;
		menu.value = 'descending';
		menu.dispatchEvent(new Event('change'));
		expect(seen).toEqual([{ type: 'sort', order: 'descending' }]);
	});
});
