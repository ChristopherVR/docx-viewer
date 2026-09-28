/** Clipboard operations for the context menu. Nothing here pretends a denied permission succeeded. */
import { deleteSelection } from 'prosemirror-commands';
import type { EditorView } from 'prosemirror-view';

function selectedText(view: EditorView): string {
	const { from, to } = view.state.selection;
	return view.state.doc.textBetween(from, to, '\n');
}

/**
 * Cut or copy through the browser's own command so ProseMirror's clipboard serializer keeps rich
 * content; falls back to plain text through the async clipboard API when the command is refused.
 */
export async function copyOrCut(view: EditorView, op: 'cut' | 'copy'): Promise<boolean> {
	view.focus();
	if (typeof document.execCommand === 'function' && document.execCommand(op)) return true;
	if (typeof navigator === 'undefined' || !navigator.clipboard?.writeText) return false;
	try {
		await navigator.clipboard.writeText(selectedText(view));
	} catch {
		return false;
	}
	if (op === 'cut' && view.editable) deleteSelection(view.state, view.dispatch);
	return true;
}

/** Whether the browser will let us read the clipboard; `true` when it cannot say. */
export async function clipboardReadAllowed(): Promise<boolean> {
	if (typeof navigator === 'undefined' || typeof navigator.clipboard?.readText !== 'function')
		return false;
	try {
		const status = await navigator.permissions?.query({ name: 'clipboard-read' as PermissionName });
		return status?.state !== 'denied';
	} catch {
		return true;
	}
}

/** Pastes clipboard text at the selection. Returns false when the browser refuses the read. */
export async function pasteText(view: EditorView): Promise<boolean> {
	try {
		const text = await navigator.clipboard.readText();
		view.focus();
		return text ? view.pasteText(text) : true;
	} catch {
		return false;
	}
}
