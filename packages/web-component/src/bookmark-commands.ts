import { closeHistory } from 'prosemirror-history';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export const MAX_BOOKMARK_LENGTH = 40;

export type BookmarkNameProblem = 'empty' | 'start' | 'characters' | 'length';

/**
 * Word's bookmark naming rule: a letter first, then letters, digits and underscores, at most 40
 * characters, no spaces. Returns the first rule the name breaks, or null when it is valid.
 */
export function bookmarkNameProblem(name: string): BookmarkNameProblem | null {
	if (!name) return 'empty';
	if (!/^\p{L}/u.test(name)) return 'start';
	if (!/^[\p{L}\p{N}_]+$/u.test(name)) return 'characters';
	if ([...name].length > MAX_BOOKMARK_LENGTH) return 'length';
	return null;
}

interface Located {
	name: string;
	pos: number;
	paragraph: number;
}

/** User bookmarks (not Word's hidden `_` ones) with the paragraph that holds each. */
function locate(view: EditorView): Located[] {
	const found: Located[] = [];
	view.state.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		for (const name of (node.attrs.bookmarks as string[] | undefined) ?? [])
			if (!name.startsWith('_')) found.push({ name, pos: pos + 1, paragraph: pos });
		return false;
	});
	return found;
}

const sameName = (a: string, b: string) => a.toLocaleLowerCase() === b.toLocaleLowerCase();

/** Bookmark names in document order. */
export function listBookmarks(view: EditorView): string[] {
	return locate(view).map((item) => item.name);
}

/**
 * Word's Bookmark > Add: puts `name` on the paragraph holding the selection. A bookmark of the
 * same name (compared without regard to case) elsewhere moves here, as in Word. Returns
 * `added`, `moved`, or the naming problem.
 */
export function addBookmark(
	view: EditorView,
	name: string,
): 'added' | 'moved' | BookmarkNameProblem | 'unavailable' {
	if (!view.editable) return 'unavailable';
	const problem = bookmarkNameProblem(name);
	if (problem) return problem;
	const here = view.state.selection.$from;
	const target = here.parent.type.name === 'paragraph' ? here.before() : undefined;
	if (target === undefined) return 'unavailable';
	const existing = locate(view).find((item) => sameName(item.name, name));
	let tr = view.state.tr;
	if (existing) {
		const old = tr.doc.nodeAt(existing.paragraph);
		if (old)
			tr = tr.setNodeAttribute(
				existing.paragraph,
				'bookmarks',
				(old.attrs.bookmarks as string[]).filter((item) => !sameName(item, name)),
			);
	}
	const node = tr.doc.nodeAt(target);
	if (!node) return 'unavailable';
	tr = tr.setNodeAttribute(target, 'bookmarks', [
		...((node.attrs.bookmarks as string[]) ?? []).filter((item) => !sameName(item, name)),
		name,
	]);
	view.dispatch(closeHistory(tr));
	return existing ? 'moved' : 'added';
}

/** Bookmark > Delete: removes the named bookmark. Returns whether it existed. */
export function deleteBookmark(view: EditorView, name: string): boolean {
	if (!view.editable) return false;
	const existing = locate(view).find((item) => item.name === name);
	if (!existing) return false;
	const node = view.state.doc.nodeAt(existing.paragraph)!;
	view.dispatch(
		closeHistory(
			view.state.tr.setNodeAttribute(
				existing.paragraph,
				'bookmarks',
				(node.attrs.bookmarks as string[]).filter((item) => item !== name),
			),
		),
	);
	return true;
}

/** Bookmark > Go To: puts the caret at the bookmark's paragraph and scrolls to it. */
export function goToBookmark(view: EditorView, name: string): boolean {
	const existing = locate(view).find((item) => item.name === name);
	if (!existing) return false;
	view.dispatch(
		view.state.tr
			.setSelection(TextSelection.near(view.state.doc.resolve(existing.pos)))
			.scrollIntoView(),
	);
	view.focus();
	return true;
}
