import type { EditorView } from 'prosemirror-view';
import type { Mark } from 'prosemirror-model';
import { schema } from './schema';
import { isNavigableHref } from './inline-content-schema';

export interface LinkTarget {
	href?: string;
	anchor?: string;
	tooltip?: string;
}

/** The contiguous range carrying the same link mark around `pos`, if the position is inside a link. */
export function linkRangeAt(
	view: EditorView,
	pos: number,
): { from: number; to: number; mark: Mark } | undefined {
	const $pos = view.state.doc.resolve(pos);
	const items: { from: number; to: number; link?: Mark }[] = [];
	let cursor = $pos.start();
	$pos.parent.forEach((child) => {
		const link = child.marks.find((item) => item.type === schema.marks.link);
		items.push({ from: cursor, to: cursor + child.nodeSize, ...(link && { link }) });
		cursor += child.nodeSize;
	});
	let index = items.findIndex((item) => item.link && item.from <= pos && pos < item.to);
	if (index < 0) index = items.findIndex((item) => item.link && item.from < pos && pos <= item.to);
	if (index < 0) return undefined;
	const mark = items[index]?.link;
	if (!mark) return undefined;
	let start = index;
	let end = index;
	while (start > 0 && items[start - 1]?.link?.eq(mark)) start--;
	while (end < items.length - 1 && items[end + 1]?.link?.eq(mark)) end++;
	const first = items[start];
	const last = items[end];
	return first && last ? { from: first.from, to: last.to, mark } : undefined;
}

/** The link under the selection start, used to prefill the link dialog. */
export function linkAtSelection(view: EditorView): LinkTarget | undefined {
	const range = linkRangeAt(view, view.state.selection.from);
	if (!range) return undefined;
	const { href, anchor, tooltip } = range.mark.attrs;
	return {
		...(href ? { href } : {}),
		...(anchor ? { anchor } : {}),
		...(tooltip ? { tooltip } : {}),
	};
}

/**
 * Applies a link to the selection. With an empty selection inside a link the whole link is
 * retargeted; with an empty selection elsewhere `displayText` (or the address) is inserted as a link.
 */
export function applyLink(view: EditorView, target: LinkTarget, displayText?: string): boolean {
	if (!target.href && !target.anchor) return false;
	if (target.href && !isNavigableHref(target.href))
		throw new Error('Links must start with http://, https:// or mailto:.');
	const mark = schema.marks.link.create({
		href: target.href ?? null,
		anchor: target.anchor ?? null,
		tooltip: target.tooltip || null,
	});
	const { state } = view;
	const { from, to, empty } = state.selection;
	const tr = state.tr;
	if (!empty) tr.removeMark(from, to, schema.marks.link).addMark(from, to, mark);
	else {
		const existing = linkRangeAt(view, from);
		if (existing)
			tr.removeMark(existing.from, existing.to, schema.marks.link).addMark(
				existing.from,
				existing.to,
				mark,
			);
		else {
			const text = displayText || target.href || target.anchor!;
			tr.insert(
				from,
				schema.text(text, [...(state.storedMarks ?? state.selection.$from.marks()), mark]),
			);
		}
	}
	view.dispatch(tr.scrollIntoView());
	return true;
}

/** Removes the link from the selection, or from the whole link under an empty selection. */
export function removeLink(view: EditorView): boolean {
	const { from, to, empty } = view.state.selection;
	const range = empty ? linkRangeAt(view, from) : { from, to };
	if (!range) return false;
	view.dispatch(view.state.tr.removeMark(range.from, range.to, schema.marks.link));
	return true;
}

/** Scrolls to the paragraph whose bookmarks include `name`; returns whether one was found. */
export function goToBookmark(view: EditorView, name: string): boolean {
	let target = -1;
	view.state.doc.descendants((node, pos) => {
		if (target >= 0) return false;
		if (
			node.type.name === 'paragraph' &&
			(node.attrs.bookmarks as string[] | undefined)?.includes(name)
		)
			target = pos;
		return true;
	});
	if (target < 0) return false;
	const dom = view.nodeDOM(target);
	if (dom instanceof HTMLElement) dom.scrollIntoView({ block: 'center' });
	return true;
}

/** Bookmark names in document order, for the "Place in this document" list. */
export function bookmarkNames(view: EditorView): string[] {
	const names: string[] = [];
	view.state.doc.descendants((node) => {
		if (node.type.name === 'paragraph') names.push(...((node.attrs.bookmarks as string[]) ?? []));
		return node.type.name !== 'paragraph';
	});
	return names.filter((name) => !name.startsWith('_'));
}

/**
 * Follows the link at `pos` like Word's Ctrl+Click: web and mail links open in a new tab,
 * internal links scroll to their bookmark. Returns whether a link was followed.
 */
export function followLinkAt(view: EditorView, pos: number): boolean {
	const range = linkRangeAt(view, pos);
	if (!range) return false;
	const { href, anchor } = range.mark.attrs;
	if (isNavigableHref(href)) {
		window.open(href, '_blank', 'noopener,noreferrer');
		return true;
	}
	return anchor ? goToBookmark(view, String(anchor)) : false;
}
