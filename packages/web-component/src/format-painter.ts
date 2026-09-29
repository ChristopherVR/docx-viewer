import type { Mark } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

interface Painter {
	marks: readonly Mark[];
	stop: () => void;
}

const armed = new WeakMap<EditorView, Painter>();

/** Character marks Word's Format Painter carries: everything except links, comments and revisions. */
const COPIED = new Set(
	['bold', 'italic', 'underline', 'strike', 'verticalAlign', 'highlight', 'font', 'characterStyle']
		.map((name) => schema.marks[name])
		.filter((type) => type !== undefined),
);

export function isFormatPainterArmed(view: EditorView): boolean {
	return armed.has(view);
}

function sourceMarks(view: EditorView): Mark[] {
	const { selection, storedMarks } = view.state;
	const marks = selection.empty
		? (storedMarks ?? selection.$from.marks())
		: (view.state.doc.nodeAt(selection.from)?.marks ?? selection.$from.marks());
	return marks.filter((mark) => COPIED.has(mark.type));
}

/** Replaces the character formatting of `[from, to]` with `marks`. */
export function paintMarks(
	view: EditorView,
	marks: readonly Mark[],
	from: number,
	to: number,
): void {
	let tr = view.state.tr;
	for (const type of COPIED) tr = tr.removeMark(from, to, type);
	for (const mark of marks) tr = tr.addMark(from, to, mark);
	view.dispatch(tr);
}

/**
 * Word's Format Painter: the first call copies the character formatting at the selection and arms
 * the painter; the next non-empty selection made with the mouse or keyboard receives it. Calling it
 * again, pressing Escape or editing cancels it. Paragraph formatting is not carried.
 * Returns whether the painter is now armed.
 */
export function toggleFormatPainter(view: EditorView, onChange: (armed: boolean) => void): boolean {
	const existing = armed.get(view);
	if (existing) {
		existing.stop();
		return false;
	}
	if (!view.editable) return false;
	const marks = sourceMarks(view);
	/** ProseMirror reads the DOM selection after mouseup/keyup, so the check waits one tick. */
	const later = () => void setTimeout(apply, 0);
	const apply = () => {
		if (!armed.has(view)) return;
		const { from, to, empty } = view.state.selection;
		if (empty) return;
		paintMarks(view, marks, from, to);
		stop();
	};
	const cancel = (event: KeyboardEvent) => {
		if (event.key === 'Escape') stop();
	};
	const stop = () => {
		view.dom.removeEventListener('mouseup', later);
		view.dom.removeEventListener('keyup', keyup);
		view.dom.removeEventListener('keydown', cancel);
		view.dom.classList.remove('dve-painting');
		armed.delete(view);
		onChange(false);
	};
	const keyup = (event: KeyboardEvent) => {
		if (event.key.startsWith('Arrow') && event.shiftKey) later();
	};
	view.dom.addEventListener('mouseup', later);
	view.dom.addEventListener('keyup', keyup);
	view.dom.addEventListener('keydown', cancel);
	view.dom.classList.add('dve-painting');
	armed.set(view, { marks, stop });
	onChange(true);
	return true;
}
