import { registerDocxEditor, type DocxEditorElement } from '@christophervr/docx-web-component';
import type { DocumentModel } from '@christophervr/docx-core';

export interface EditorOptions {
	documentModel?: DocumentModel;
	readOnly?: boolean;
	locale?: string;
	onDocumentChange?: (model: DocumentModel) => void;
	onDocumentError?: (error: Error) => void;
}
export interface EditorHandle {
	readonly element: DocxEditorElement;
	load(input: Uint8Array | ArrayBuffer): Promise<void>;
	save(): Promise<Uint8Array>;
}
export interface EditorBinding extends EditorHandle {
	update(options: EditorOptions): void;
	destroy(): void;
}
/** All framework adapters share property, event and lifecycle semantics here. */
export function mountEditor(host: HTMLElement, initial: EditorOptions = {}): EditorBinding {
	registerDocxEditor();
	const element = host.ownerDocument.createElement('docx-editor') as DocxEditorElement;
	let options: EditorOptions = {};
	let lastInput: DocumentModel | undefined;
	let lastEmitted: DocumentModel | undefined;
	let destroyed = false;
	const changed = (event: Event) => {
		lastEmitted = (event as CustomEvent<DocumentModel>).detail;
		options.onDocumentChange?.(lastEmitted);
	};
	const failed = (event: Event) => options.onDocumentError?.((event as CustomEvent<Error>).detail);
	element.addEventListener('document-change', changed);
	element.addEventListener('document-error', failed);
	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			element.locale = next.locale ?? 'en';
			element.readOnly = next.readOnly ?? false;
			if (
				next.documentModel &&
				next.documentModel !== lastInput &&
				next.documentModel !== lastEmitted
			) {
				element.documentModel = next.documentModel;
			}
			lastInput = next.documentModel;
		},
		load: (input) => element.load(input),
		save: () => element.save(),
		destroy() {
			if (destroyed) return;
			destroyed = true;
			element.removeEventListener('document-change', changed);
			element.removeEventListener('document-error', failed);
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	return binding;
}
