import type { EditorView } from 'prosemirror-view';
import { TextSelection } from 'prosemirror-state';

/**
 * Maps a Print Layout click (`{ blockId, offset }`, from `print-layout.ts`'s
 * `resolveClick`) back to a ProseMirror document position: finds the
 * paragraph node whose `id` attribute matches, then clamps `offset` into its
 * text content. Best-effort by design (see `print-layout-view.ts`); it does
 * not attempt sub-run fragment precision.
 */
export function moveCursorToBlock(view: EditorView, blockId: string, offset: number): boolean {
	let target: { pos: number; size: number } | undefined;
	view.state.doc.descendants((node, pos) => {
		if (target) return false;
		if (node.type.name === 'paragraph' && node.attrs.id === blockId) {
			target = { pos: pos + 1, size: node.content.size };
			return false;
		}
		return true;
	});
	if (!target) return false;
	const clamped = Math.max(0, Math.min(offset, target.size));
	const selection = TextSelection.near(view.state.doc.resolve(target.pos + clamped));
	view.dispatch(view.state.tr.setSelection(selection));
	view.focus();
	return true;
}
