import type { EditorView } from 'prosemirror-view';

/**
 * Word's Spelling toggle drives the browser's own spell checker on the editing surface. It does not
 * bundle a dictionary, grammar checker or Word's Editor pane.
 */
export function syncSpellingButton(toolbar: HTMLElement | undefined, view: EditorView): void {
	toolbar
		?.querySelector('[aria-label="Spelling"], [data-localearialabel="Spelling"]')
		?.setAttribute('aria-pressed', String(view.dom.spellcheck !== false));
}
