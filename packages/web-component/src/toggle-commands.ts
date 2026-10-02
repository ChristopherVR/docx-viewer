import { resolveRunFormatting, type TextRun } from 'docx-core';
import { toggleMark } from 'prosemirror-commands';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import { appendInlineNode } from './run-adapter';
import { styleModelOf } from './run-styles';
import { schema } from './schema';

export type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

interface Piece {
	from: number;
	to: number;
	node: ProseMirrorNode;
	paragraphStyle?: string;
}

function selectedText(state: EditorState): Piece[] {
	const pieces: Piece[] = [];
	const { from, to } = state.selection;
	state.doc.nodesBetween(from, to, (node, pos, parent) => {
		if (!node.isText) return true;
		pieces.push({
			from: Math.max(pos, from),
			to: Math.min(pos + node.nodeSize, to),
			node,
			paragraphStyle: parent?.attrs.style || undefined,
		});
		return false;
	});
	return pieces;
}

/** Whether a piece of text shows `key`, from its marks, explicit offs and inherited styles. */
function isOn(state: EditorState, piece: Piece, key: ToggleKey): boolean {
	const runs: TextRun[] = [];
	appendInlineNode(runs, piece.node);
	const run = runs[0];
	if (!run) return false;
	if (run[key] !== undefined) return Boolean(run[key]);
	const model = styleModelOf(state);
	if (!model) return false;
	return Boolean(
		resolveRunFormatting(run, {
			runCatalog: model.characterStyles,
			paragraphCatalog: model.paragraphStyles,
			paragraphStyleId: piece.paragraphStyle,
		})[key],
	);
}

/** Sets or clears the explicit off for `key` in a piece's opaque run-properties mark. */
function setExplicitOff(tr: Transaction, piece: Piece, key: ToggleKey, off: boolean): void {
	const type = schema.marks.runProperties;
	const current = type.isInSet(piece.node.marks);
	const props = { ...((current?.attrs.props as Record<string, unknown> | null) ?? {}) };
	if (off) props[key] = false;
	else delete props[key];
	tr.removeMark(piece.from, piece.to, type);
	if (Object.keys(props).length) tr.addMark(piece.from, piece.to, type.create({ props }));
}

/**
 * Word's Bold/Italic/Underline/Strikethrough: toggles what the text shows, including formatting
 * inherited from its styles. Turning off text that is on only through a style (a bold heading)
 * writes an explicit off, as Word does; turning it on clears that off again.
 */
export function toggleFormat(key: ToggleKey): Command {
	const markType = schema.marks[key];
	const plain = toggleMark(markType);
	return (state, dispatch) => {
		if (state.selection.empty) return plain(state, dispatch);
		const pieces = selectedText(state);
		if (!pieces.length) return plain(state, dispatch);
		const allOn = pieces.every((piece) => isOn(state, piece, key));
		if (!dispatch) return true;
		const tr = state.tr;
		for (const piece of pieces) {
			if (allOn) {
				tr.removeMark(piece.from, piece.to, markType);
				// Still on through a style once the mark is gone: cancel it explicitly.
				const withoutMark = {
					...piece,
					node: piece.node.mark(markType.removeFromSet(piece.node.marks)),
				};
				setExplicitOff(tr, withoutMark, key, false);
				const inherited = isOn(
					state,
					{
						...withoutMark,
						node: withoutMark.node.mark(
							schema.marks.runProperties.removeFromSet(withoutMark.node.marks),
						),
					},
					key,
				);
				if (inherited) setExplicitOff(tr, withoutMark, key, true);
			} else {
				setExplicitOff(tr, piece, key, false);
				tr.addMark(piece.from, piece.to, markType.create());
			}
		}
		dispatch(tr.scrollIntoView());
		return true;
	};
}
