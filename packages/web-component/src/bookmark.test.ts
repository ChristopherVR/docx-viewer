// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import {
	addBookmark,
	bookmarkNameProblem,
	deleteBookmark,
	goToBookmark,
	listBookmarks,
} from './bookmark-commands';
import { FormatDialogs } from './format-dialogs';
import { createRibbon } from './ribbon';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});
afterEach(() => (document.body.innerHTML = ''));

function editor(...paragraphs: Array<Record<string, unknown>>) {
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((attrs, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...attrs }, schema.text(`text ${index}`)),
		),
	);
	const host = document.createElement('div');
	document.body.append(host);
	return new EditorView(host, { state: EditorState.create({ doc, schema }) });
}
const caretIn = (view: EditorView, index: number) => {
	let pos = 0;
	view.state.doc.forEach((node, offset, i) => {
		if (i === index) pos = offset + 1;
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
};
const marks = (view: EditorView, index: number) =>
	view.state.doc.child(index).attrs.bookmarks as string[];

describe('bookmarkNameProblem', () => {
	it("applies Word's naming rule", () => {
		expect(bookmarkNameProblem('')).toBe('empty');
		expect(bookmarkNameProblem('1st')).toBe('start');
		expect(bookmarkNameProblem('_hidden')).toBe('start');
		expect(bookmarkNameProblem('has space')).toBe('characters');
		expect(bookmarkNameProblem('a-b')).toBe('characters');
		expect(bookmarkNameProblem('a'.repeat(41))).toBe('length');
		expect(bookmarkNameProblem('Chapter_1')).toBeNull();
		expect(bookmarkNameProblem('Überblick2')).toBeNull();
		expect(bookmarkNameProblem('a'.repeat(40))).toBeNull();
	});
});

describe('bookmark commands', () => {
	it('adds a bookmark to the caret paragraph and lists it', () => {
		const view = editor({}, {});
		caretIn(view, 1);
		expect(addBookmark(view, 'Summary')).toBe('added');
		expect(marks(view, 1)).toEqual(['Summary']);
		expect(listBookmarks(view)).toEqual(['Summary']);
	});

	it('moves a bookmark of the same name, ignoring case, instead of duplicating it', () => {
		const view = editor({ bookmarks: ['Summary'] }, {});
		caretIn(view, 1);
		expect(addBookmark(view, 'summary')).toBe('moved');
		expect(marks(view, 0)).toEqual([]);
		expect(marks(view, 1)).toEqual(['summary']);
	});

	it('rejects invalid names, read-only views and hidden bookmarks in the list', () => {
		const view = editor({ bookmarks: ['_Toc1', 'Mine'] });
		expect(addBookmark(view, 'bad name')).toBe('characters');
		expect(listBookmarks(view)).toEqual(['Mine']);
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(addBookmark(locked, 'Fine')).toBe('unavailable');
		expect(deleteBookmark(locked, 'Mine')).toBe(false);
	});

	it('deletes only the named bookmark and reports a missing one', () => {
		const view = editor({ bookmarks: ['A1', 'B1', '_Toc9'] });
		expect(deleteBookmark(view, 'A1')).toBe(true);
		expect(marks(view, 0)).toEqual(['B1', '_Toc9']);
		expect(deleteBookmark(view, 'Nope')).toBe(false);
	});

	it('goes to a bookmark', () => {
		const view = editor({}, { bookmarks: ['Target'] });
		caretIn(view, 0);
		expect(goToBookmark(view, 'Target')).toBe(true);
		expect(view.state.selection.$from.parent.textContent).toBe('text 1');
		expect(goToBookmark(view, 'Missing')).toBe(false);
	});
});

describe('Bookmark dialog', () => {
	const setup = (...attrs: Array<Record<string, unknown>>) => {
		const view = editor(...attrs);
		const dialogs = new FormatDialogs({ view: () => view, model: () => createDocument() });
		document.body.append(...dialogs.elements);
		dialogs.open('bookmark');
		return { view, root: dialogs.elements[2]!, dialogs };
	};
	const field = (root: ParentNode, label: string) =>
		root.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${label}"]`)!;
	const button = (root: ParentNode, text: string) =>
		[...root.querySelectorAll<HTMLButtonElement>('button')].find((b) => b.textContent === text)!;

	it('adds a bookmark by name and closes', () => {
		const { view, root } = setup({});
		expect(button(root, 'Add').disabled).toBe(true);
		const name = field(root, 'Bookmark name');
		name.value = 'Intro';
		name.dispatchEvent(new Event('input'));
		expect(button(root, 'Add').disabled).toBe(false);
		button(root, 'Add').click();
		expect(root.hidden).toBe(true);
		expect(marks(view, 0)).toEqual(['Intro']);
	});

	it('explains an invalid name and keeps Add disabled', () => {
		const { root } = setup({});
		const name = field(root, 'Bookmark name');
		name.value = '9lives';
		name.dispatchEvent(new Event('input'));
		expect(root.querySelector('.dve-dialog-message')!.textContent).toBe(
			'A bookmark name must begin with a letter.',
		);
		expect(button(root, 'Add').disabled).toBe(true);
	});

	it('lists existing bookmarks, selects one into the name, and deletes it', () => {
		const { view, root } = setup({ bookmarks: ['Beta', 'Alpha'] });
		const list = field(root, 'Bookmarks');
		expect([...list.options].map((o) => o.value)).toEqual(['Alpha', 'Beta']);
		expect(button(root, 'Delete').disabled).toBe(true);
		list.value = 'Beta';
		list.dispatchEvent(new Event('change'));
		expect(field(root, 'Bookmark name').value).toBe('Beta');
		expect(button(root, 'Delete').disabled).toBe(false);
		button(root, 'Delete').click();
		expect(marks(view, 0)).toEqual(['Alpha']);
		expect(root.hidden).toBe(false);
	});

	it('keeps no fields in the document while closed', () => {
		const { dialogs } = setup({});
		dialogs.closeAll();
		expect(dialogs.elements[2]!.childElementCount).toBe(0);
	});

	it('has a ribbon button on the Insert tab', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Bookmark"]')!.click();
		expect(seen).toEqual([{ type: 'formatDialog', kind: 'bookmark' }]);
	});
});
