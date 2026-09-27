import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorLocale } from './localization';

/** What feature controllers need from the editor element, without reaching into its internals. */
export interface EditorHost {
	readonly element: HTMLElement;
	view(): EditorView | undefined;
	model(): DocumentModel;
	/** Replaces the model for edits made outside the main editing surface. */
	setModel(model: DocumentModel): void;
	locale(): EditorLocale;
	/** Editable and not in a collaboration session (parts and page setup aren't synced). */
	canEditOutsideBody(): boolean;
	/** Marks the document changed after an edit outside the main surface. */
	edited(): void;
	reportError(cause: unknown): void;
}

/** Raises the editor's `document-error` event for any thrown value. */
export function dispatchDocumentError(element: HTMLElement, cause: unknown): void {
	element.dispatchEvent(
		new CustomEvent('document-error', {
			detail: cause instanceof Error ? cause : new Error(String(cause)),
			bubbles: true,
			composed: true,
		}),
	);
}
