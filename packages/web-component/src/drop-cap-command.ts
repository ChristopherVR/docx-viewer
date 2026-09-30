import { closeHistory } from 'prosemirror-history';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { effectiveFont } from './font-sync';
import { schema } from './schema';

export type DropCapStyle = 'none' | 'drop' | 'margin';

const LINES = 3;

/** The drop cap frame paragraph and the paragraph it belongs to, around the caret. */
function findPair(view: EditorView): { framePos?: number; bodyPos: number } | undefined {
	const { $from } = view.state.selection;
	if ($from.parent.type.name !== 'paragraph') return undefined;
	const index = $from.index($from.depth - 1);
	const container = $from.node($from.depth - 1);
	const here = $from.before();
	if ($from.parent.attrs.dropCap) {
		const body = container.maybeChild(index + 1);
		return body?.type.name === 'paragraph'
			? { framePos: here, bodyPos: here + $from.parent.nodeSize }
			: undefined;
	}
	const previous = index > 0 ? container.child(index - 1) : undefined;
	if (previous?.attrs.dropCap) return { framePos: here - previous.nodeSize, bodyPos: here };
	return { bodyPos: here };
}

/** Marks for the enlarged initial: its own marks with the font size set for `lines` lines. */
function capMarks(node: ProseMirrorNode, size: number) {
	const marks = node.marks.filter((mark) => mark.type !== schema.marks.font);
	const font = node.marks.find((mark) => mark.type === schema.marks.font);
	return [...marks, schema.marks.font.create({ ...font?.attrs, size })];
}

/** Marks without an enlarged size, for folding the initial back into its paragraph. */
function plainMarks(node: ProseMirrorNode) {
	return node.marks.flatMap((mark) => {
		if (mark.type !== schema.marks.font) return [mark];
		return mark.attrs.family || mark.attrs.color
			? [schema.marks.font.create({ ...mark.attrs, size: null })]
			: [];
	});
}

/**
 * Insert > Drop Cap. Like Word, the first letter of the paragraph moves into its own frame
 * paragraph (`w:framePr`) placed before it; None folds the letter back. Choosing a style on a
 * paragraph that already has a frame only restyles the frame.
 */
export function setDropCap(view: EditorView, style: DropCapStyle): boolean {
	if (!view.editable) return false;
	const pair = findPair(view);
	if (!pair) return false;
	const { doc } = view.state;
	const body = doc.nodeAt(pair.bodyPos)!;
	const frame = pair.framePos === undefined ? undefined : doc.nodeAt(pair.framePos)!;
	let tr = view.state.tr;

	if (style === 'none') {
		if (!frame) return false;
		const letters: ProseMirrorNode[] = [];
		frame.forEach((child) =>
			letters.push(child.isText ? schema.text(child.text!, plainMarks(child)) : child),
		);
		tr = tr
			.insert(pair.bodyPos + 1, letters)
			.delete(pair.framePos!, pair.framePos! + frame.nodeSize);
		tr.setSelection(TextSelection.near(tr.doc.resolve(pair.framePos! + 1)));
		view.dispatch(closeHistory(tr).scrollIntoView());
		return true;
	}

	const cap = { style, lines: LINES };
	if (frame) {
		tr = tr.setNodeMarkup(pair.framePos!, undefined, { ...frame.attrs, dropCap: cap });
		view.dispatch(closeHistory(tr));
		return true;
	}
	const letter = body.firstChild;
	if (!letter?.isText || !letter.text || /^\s/.test(letter.text)) return false;
	const first = letter.text.codePointAt(0)!;
	const char = String.fromCodePoint(first);
	const base = effectiveFont(view.state, letter, body).size;
	const size = Math.round(base * LINES * 1.65);
	const initial = schema.text(char, capMarks(letter, size));
	const frameNode = schema.nodes.paragraph!.create(
		{
			dropCap: cap,
			lineSpacingRule: 'exact',
			lineSpacingTwips: Math.round(base * LINES * 1.2 * 20),
			spacingBeforeTwips: 0,
			spacingAfterTwips: 0,
		},
		initial,
	);
	tr = tr.delete(pair.bodyPos + 1, pair.bodyPos + 1 + char.length).insert(pair.bodyPos, frameNode);
	tr.setSelection(TextSelection.near(tr.doc.resolve(pair.bodyPos + frameNode.nodeSize + 1)));
	view.dispatch(closeHistory(tr).scrollIntoView());
	return true;
}
