// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { beforeAll, describe, expect, it } from 'vitest';
import { commentAnchors, commentIdsAtSelection, goToComment } from './comment-commands';
import { createRibbon } from './ribbon';
import { schema } from './schema';

beforeAll(() => {
	const rects = { length: 0, item: () => null, [Symbol.iterator]: function* () {} };
	Object.assign(Range.prototype, {
		getClientRects: () => rects,
		getBoundingClientRect: () => ({ top: 0, left: 0, right: 0, bottom: 0, width: 0, height: 0 }),
	});
});

/** "aaa BBB ccc DDD" with comments c1 on BBB and c2 on DDD (overlapping c3 on the tail of DDD). */
function editor() {
	const mark = (...ids: string[]) => [schema.marks.comment!.create({ ids })];
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p' }, [
			schema.text('aaa '),
			schema.text('BBB', mark('c1')),
			schema.text(' ccc '),
			schema.text('DD', mark('c2')),
			schema.text('D', mark('c2', 'c3')),
		]),
	]);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const caret = (view: EditorView, pos: number) =>
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));

describe('comment anchors', () => {
	it('lists comments by where their text starts', () => {
		expect(commentAnchors(editor())).toEqual([
			{ id: 'c1', from: 5 },
			{ id: 'c2', from: 13 },
			{ id: 'c3', from: 15 },
		]);
	});

	it('finds the comments at the caret or selection, outermost first', () => {
		const view = editor();
		caret(view, 3);
		expect(commentIdsAtSelection(view)).toEqual([]);
		caret(view, 7);
		expect(commentIdsAtSelection(view)).toEqual(['c1']);
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 13, 16)));
		expect(commentIdsAtSelection(view)).toEqual(['c2', 'c3']);
	});
});

describe('goToComment', () => {
	it('moves forward and wraps to the first comment', () => {
		const view = editor();
		caret(view, 1);
		expect(goToComment(view, 'next')).toBe('c1');
		expect(view.state.selection.from).toBe(5);
		expect(goToComment(view, 'next')).toBe('c2');
		expect(goToComment(view, 'next')).toBe('c3');
		expect(goToComment(view, 'next')).toBe('c1');
	});

	it('moves backward and wraps to the last comment', () => {
		const view = editor();
		caret(view, 6);
		expect(goToComment(view, 'previous')).toBe('c1');
		expect(goToComment(view, 'previous')).toBe('c3');
		expect(goToComment(view, 'previous')).toBe('c2');
	});

	it('does nothing without comments', () => {
		const doc = schema.node('doc', null, [
			schema.nodes.paragraph!.create({ id: 'p' }, schema.text('x')),
		]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		expect(goToComment(view, 'next')).toBeNull();
	});
});

describe('Review comment buttons', () => {
	it('emit delete, previous and next actions', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		for (const name of ['Delete comment', 'Previous comment', 'Next comment'])
			ribbon.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!.click();
		expect(seen).toEqual([
			{ type: 'comments', key: 'delete' },
			{ type: 'comments', key: 'previous' },
			{ type: 'comments', key: 'next' },
		]);
	});
});
