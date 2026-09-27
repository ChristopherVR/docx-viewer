import type { EditorView } from 'prosemirror-view';

/** Returns focus to the editor where the platform supports it (jsdom and SSR do not). */
export function focusView(view: EditorView | undefined): void {
	if (view && typeof document !== 'undefined' && typeof document.execCommand === 'function')
		view.focus();
}
