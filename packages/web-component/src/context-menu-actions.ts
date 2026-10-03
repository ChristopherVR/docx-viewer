/** Clipboard operations for the context menu. Nothing here pretends a denied permission succeeded. */
import { deleteSelection } from 'prosemirror-commands';
import type { EditorView } from 'prosemirror-view';

/**
 * Cut or copy through the browser's own command so ProseMirror's clipboard serializer keeps rich
 * content; the async fallback keeps its HTML slice when that API is available.
 */
export async function copyOrCut(view: EditorView, op: 'cut' | 'copy'): Promise<boolean> {
	if (view.state.selection.empty || (op === 'cut' && !view.editable)) return false;
	view.focus();
	if (typeof document.execCommand === 'function' && document.execCommand(op)) return true;
	if (typeof navigator === 'undefined' || !navigator.clipboard) return false;
	const { doc, selection } = view.state;
	const { dom, text } = view.serializeForClipboard(selection.content());
	try {
		if (navigator.clipboard.write && typeof ClipboardItem !== 'undefined') {
			await navigator.clipboard.write([
				new ClipboardItem({
					'text/html': new Blob([dom.innerHTML], { type: 'text/html' }),
					'text/plain': new Blob([text], { type: 'text/plain' }),
				}),
			]);
		} else if (navigator.clipboard.writeText) await navigator.clipboard.writeText(text);
		else return false;
	} catch {
		return false;
	}
	// A permission prompt can outlive the selection: never delete newly selected content.
	if (
		op === 'cut' &&
		!view.isDestroyed &&
		view.editable &&
		view.state.doc.eq(doc) &&
		view.state.selection.eq(selection)
	)
		deleteSelection(view.state, (transaction) =>
			view.dispatch(transaction.setMeta('uiEvent', 'cut')),
		);
	return true;
}

/** Whether the browser will let us read the clipboard; `true` when it cannot say. */
export async function clipboardReadAllowed(): Promise<boolean> {
	if (
		typeof navigator === 'undefined' ||
		(!navigator.clipboard?.read && !navigator.clipboard?.readText)
	)
		return false;
	try {
		const status = await navigator.permissions?.query({ name: 'clipboard-read' as PermissionName });
		return status?.state !== 'denied';
	} catch {
		return true;
	}
}

/** Uses ProseMirror's HTML parser for formatted content, with a plain-text API fallback. */
export async function pasteClipboard(view: EditorView): Promise<boolean> {
	if (!view.editable) return false;
	try {
		if (navigator.clipboard.read) {
			const items = await navigator.clipboard.read();
			for (const item of items) {
				if (!item.types.includes('text/html')) continue;
				const html = await (await item.getType('text/html')).text();
				if (!view.editable || view.isDestroyed) return false;
				view.focus();
				return html ? view.pasteHTML(html) : true;
			}
			for (const item of items) {
				if (!item.types.includes('text/plain')) continue;
				const text = await (await item.getType('text/plain')).text();
				if (!view.editable || view.isDestroyed) return false;
				view.focus();
				return text ? view.pasteText(text) : true;
			}
			return false;
		}
		const text = await navigator.clipboard.readText();
		if (!view.editable || view.isDestroyed) return false;
		view.focus();
		return text ? view.pasteText(text) : true;
	} catch {
		return false;
	}
}
