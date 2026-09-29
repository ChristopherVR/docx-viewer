import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { applyFont } from './ribbon-commands';

/** Word's Grow Font / Shrink Font ladder, in points. */
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
const DEFAULT_SIZE = 11;

/** The next size on Word's ladder above or below `size`; beyond the ends it steps by 2 pt. */
export function steppedFontSize(size: number, direction: 'grow' | 'shrink'): number {
	if (direction === 'grow') return SIZES.find((step) => step > size) ?? Math.min(size + 2, 1638);
	const smaller = SIZES.filter((step) => step < size).pop();
	return smaller ?? Math.max(size - 2, 1);
}

/** Grows or shrinks the font size of the selection (or of the next typed text) by one step. */
export function stepFontSize(view: EditorView, direction: 'grow' | 'shrink'): void {
	const { state } = view;
	const marks = state.selection.empty
		? (state.storedMarks ?? state.selection.$from.marks())
		: (state.doc.nodeAt(state.selection.from)?.marks ?? state.selection.$from.marks());
	const size = Number(marks.find((mark) => mark.type === schema.marks.font)?.attrs.size);
	applyFont(view, 'size', String(steppedFontSize(size || DEFAULT_SIZE, direction)));
}
