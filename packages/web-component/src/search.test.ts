// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { history, redo, undo } from 'prosemirror-history';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { SearchController, findTextMatches } from './search-controller';
import { createSearchPanel } from './search-panel';

const bold = schema.marks.bold.create();
const italic = schema.marks.italic.create();
const paragraph = (id: string, ...children: ReturnType<typeof schema.text>[]) =>
	schema.nodes.paragraph.create({ id }, children);

function makeView(doc: ReturnType<typeof schema.node>) {
	const host = document.createElement('div');
	document.body.append(host);
	let view: EditorView;
	const state = EditorState.create({ doc, plugins: [history()] });
	view = new EditorView(host, {
		state,
		dispatchTransaction(transaction) {
			view.updateState(view.state.apply(transaction));
		},
	});
	return view;
}

afterEach(() => document.body.replaceChildren());

describe('SearchController', () => {
	it('finds across adjacent differently marked runs and replaces without disturbing surrounding marks or paragraph ids', () => {
		const doc = schema.node('doc', null, [
			paragraph(
				'stable-id',
				schema.text('Hel', [bold]),
				schema.text('lo ', [italic]),
				schema.text('HELLO'),
			),
		]);
		const view = makeView(doc);
		const search = new SearchController(() => view);
		expect(search.search('hello').count).toBe(2);
		const first = search.findNext()!;
		const second = search.findNext()!;
		expect(second.from).toBeGreaterThan(first.from);
		expect(search.findPrevious()).toEqual(first);
		expect(search.replaceOne('$&😀')).toBe(true);
		expect(view.state.doc.textContent).toBe('$&😀 HELLO');
		expect(view.state.doc.firstChild?.attrs.id).toBe('stable-id');
		const replacement = view.state.doc.firstChild!.firstChild!;
		expect(replacement.text).toBe('$&😀');
		expect(replacement.marks.map((mark) => mark.type.name)).toContain('bold');
		const remaining = view.state.doc.firstChild!.child(1);
		expect(remaining.text).toBe(' ');
		expect(remaining.marks.map((mark) => mark.type.name)).toContain('italic');
		view.destroy();
	});

	it('uses Unicode case folding without matching half of a surrogate pair or a combining grapheme', () => {
		const text = 'A😀e\u0301 CAFÉ';
		const doc = schema.node('doc', null, paragraph('unicode', schema.text(text)));
		expect(findTextMatches(doc, '😀')).toHaveLength(1);
		expect(findTextMatches(doc, '\ud83d')).toHaveLength(0);
		expect(findTextMatches(doc, 'e')).toHaveLength(0);
		expect(findTextMatches(doc, 'e\u0301')).toHaveLength(1);
		expect(findTextMatches(doc, 'café')).toHaveLength(1);
		expect(findTextMatches(doc, 'café', true)).toHaveLength(0);
	});

	it('does not match across paragraphs, cells, or hard breaks', () => {
		const lineOne = paragraph('one', schema.text('foo'), schema.text('bar', [bold]));
		const lineTwo = paragraph(
			'two',
			schema.text('fo'),
			schema.nodes.hardBreak.create(),
			schema.text('obar'),
		);
		const table = schema.nodes.table.create(null, [
			schema.nodes.tableRow.create(null, [
				schema.nodes.tableCell.create(null, paragraph('left-cell', schema.text('left'))),
				schema.nodes.tableCell.create(null, paragraph('right-cell', schema.text('right'))),
			]),
		]);
		const doc = schema.node('doc', null, [lineOne, lineTwo, table]);
		expect(findTextMatches(doc, 'foobar')).toHaveLength(1);
		expect(findTextMatches(doc, 'fo\nobar')).toHaveLength(0);
		expect(findTextMatches(doc, 'leftright')).toHaveLength(0);
		expect(findTextMatches(doc, 'left')).toHaveLength(1);
	});

	it('replaces all matches in one undoable transaction and emits hard breaks for newline replacement', () => {
		const original = schema.node('doc', null, [
			paragraph('first', schema.text('cat ', [bold]), schema.text('cat', [italic])),
			paragraph('second', schema.text('cat')),
		]);
		const view = makeView(original);
		const search = new SearchController(() => view);
		search.search('cat', true);
		expect(search.replaceAll('dog\ncat')).toBe(3);
		expect(view.state.doc.child(0).textContent).toBe('dog\ncat dog\ncat');
		expect(view.state.doc.child(1).textContent).toBe('dog\ncat');
		const breaks: number[] = [];
		view.state.doc.descendants((node, pos) => {
			if (node.type.name === 'hardBreak') breaks.push(pos);
		});
		expect(breaks).toHaveLength(3);
		const replacedNodes: Array<{ text?: string; marks: string[] }> = [];
		view.state.doc.descendants((node) => {
			if (node.isText && node.text === 'dog')
				replacedNodes.push({ text: node.text, marks: node.marks.map((mark) => mark.type.name) });
		});
		expect(replacedNodes.map((node) => node.marks)).toEqual([['bold'], ['italic'], []]);
		expect(view.state.doc.child(0).attrs.id).toBe('first');
		expect(view.state.doc.child(1).attrs.id).toBe('second');
		undo(view.state, (transaction) => view.dispatch(transaction), view);
		expect(view.state.doc.eq(original)).toBe(true);
		redo(view.state, (transaction) => view.dispatch(transaction), view);
		expect(view.state.doc.child(0).textContent).toBe('dog\ncat dog\ncat');
		expect(view.state.doc.child(1).textContent).toBe('dog\ncat');
		view.setProps({ editable: () => false });
		const unchanged = view.state.doc;
		expect(search.replaceAll('blocked')).toBe(0);
		expect(view.state.doc.eq(unchanged)).toBe(true);
		view.destroy();
	});

	it('treats empty queries safely and supports delete replacement as a literal operation', () => {
		const view = makeView(schema.node('doc', null, paragraph('empty', schema.text('a$a a$a'))));
		const search = new SearchController(() => view);
		expect(search.search('').count).toBe(0);
		expect(search.replaceAll('x')).toBe(0);
		search.search('$a', true);
		expect(search.replaceOne('')).toBe(true);
		expect(view.state.doc.textContent).toBe('a a$a');
		view.destroy();
	});
});

describe('accessible search panel', () => {
	it('exposes search controls, updates its count, disables replacement in read-only mode, and closes on Escape', () => {
		const view = makeView(schema.node('doc', null, paragraph('panel', schema.text('one one'))));
		const onClose = vi.fn();
		const panel = createSearchPanel({ getView: () => view, onClose });
		document.body.append(panel.element);
		panel.open();
		const query = panel.element.querySelector<HTMLInputElement>('[aria-label="Find text"]')!;
		query.value = 'one';
		query.dispatchEvent(new Event('input', { bubbles: true }));
		query.focus();
		query.setSelectionRange(1, 2);
		panel.refresh();
		expect(query.selectionStart).toBe(1);
		expect(query.selectionEnd).toBe(2);
		expect(panel.element.querySelector('[role="status"]')?.textContent).toBe('2 matches');
		for (const label of ['Find previous', 'Find next', 'Replace', 'Replace all', 'Close search'])
			expect(panel.element.querySelector(`[aria-label="${label}"]`)).not.toBeNull();
		expect(panel.element.querySelector('[aria-label="Match case"]')).not.toBeNull();
		view.setProps({ editable: () => false });
		panel.refresh();
		expect(
			panel.element.querySelector<HTMLButtonElement>('[aria-label="Replace all"]')?.disabled,
		).toBe(true);
		query.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(panel.isOpen).toBe(false);
		expect(onClose).toHaveBeenCalledOnce();
		view.destroy();
	});
});
