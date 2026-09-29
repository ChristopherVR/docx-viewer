import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export type SortOrder = 'ascending' | 'descending';

/**
 * Home > Sort: orders the selected paragraphs by their text, using the display language's
 * collation with numbers compared as numbers ("Item 2" before "Item 10"). Only top-level
 * paragraphs sort; a selection that includes a table is refused rather than partly sorted.
 * Returns whether anything moved.
 */
export function sortParagraphs(view: EditorView, order: SortOrder, locale?: string): boolean {
	if (!view.editable) return false;
	const { doc, selection } = view.state;
	const items: ProseMirrorNode[] = [];
	let from = -1;
	let to = -1;
	let unsortable = false;
	doc.forEach((node, offset) => {
		const end = offset + node.nodeSize;
		if (end <= selection.from || offset >= selection.to) return;
		if (node.type.name !== 'paragraph') {
			unsortable = true;
			return;
		}
		if (from < 0) from = offset;
		to = end;
		items.push(node);
	});
	if (unsortable || items.length < 2) return false;
	const collator = new Intl.Collator(locale, { numeric: true, sensitivity: 'base' });
	const sorted = [...items].sort(
		(a, b) => collator.compare(a.textContent, b.textContent) * (order === 'ascending' ? 1 : -1),
	);
	if (sorted.every((node, index) => node === items[index])) return false;
	const tr = view.state.tr.replaceWith(from, to, sorted);
	// The sorted range stays selected, as in Word, so the order can be flipped straight away.
	view.dispatch(
		closeHistory(
			tr.setSelection(TextSelection.create(tr.doc, from + 1, Math.max(from + 1, to - 1))),
		),
	);
	return true;
}
