import {
	registerDocxEditor,
	type DocxEditorElement,
	type DocxEditorEventDetail,
	type DocxEditorEventName,
	type EditorThemeMode,
} from '@christophervr/docx-web-component';
import type { DocumentModel } from '@christophervr/docx-core';

/** Editor state a framework passes down as props. */
export interface EditorProps {
	documentModel?: DocumentModel;
	readOnly?: boolean;
	locale?: string;
	/** `light`, `dark`, or `auto` (default) to follow the OS color scheme. */
	theme?: EditorThemeMode;
}
/** Editor callbacks; each is the framework-neutral form of one entry in `EDITOR_EVENT_NAMES`. */
export interface EditorEventOptions {
	onDocumentChange?: (model: DocumentModel) => void;
	onDocumentError?: (error: Error) => void;
}
export interface EditorOptions extends EditorProps, EditorEventOptions {}

/** Single source of truth for the prop keys every adapter forwards. */
export const EDITOR_PROP_KEYS = [
	'documentModel',
	'readOnly',
	'locale',
	'theme',
] as const satisfies readonly (keyof EditorProps)[];
export type EditorPropKey = (typeof EDITOR_PROP_KEYS)[number];
// Compile-time guard: adding a key to EditorProps without listing it above is an error.
const propKeysAreComplete: Exclude<keyof EditorProps, EditorPropKey> extends never ? true : never =
	true;
void propKeysAreComplete;

/** Single source of truth for the element events every adapter surfaces. */
export const EDITOR_EVENT_NAMES = [
	'document-change',
	'document-error',
] as const satisfies readonly DocxEditorEventName[];
export type EditorEventName = (typeof EDITOR_EVENT_NAMES)[number];
/** One handler per bound event; a missing key is a compile error in every adapter. */
export type EditorEventHandlers = {
	[K in EditorEventName]: ((detail: DocxEditorEventDetail<K>) => void) | undefined;
};

function copyProp<K extends EditorPropKey>(to: EditorProps, from: EditorProps, key: K): void {
	to[key] = from[key];
}
/** Copies exactly the shared props out of a framework's props/instance object. */
export function pickEditorProps(source: EditorProps): EditorProps {
	const picked: EditorProps = {};
	for (const key of EDITOR_PROP_KEYS) copyProp(picked, source, key);
	return picked;
}
/** Maps per-event handlers (keyed by DOM event name) onto the shared option callbacks. */
export function eventOptions(handlers: EditorEventHandlers): EditorEventOptions {
	return {
		onDocumentChange: handlers['document-change'],
		onDocumentError: handlers['document-error'],
	};
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
	const element = host.ownerDocument.createElement('docx-editor');
	let options: EditorOptions = {};
	let lastInput: DocumentModel | undefined;
	let lastEmitted: DocumentModel | undefined;
	let destroyed = false;
	const changed = (event: CustomEvent<DocumentModel>) => {
		lastEmitted = event.detail;
		options.onDocumentChange?.(lastEmitted);
	};
	const failed = (event: CustomEvent<Error>) => options.onDocumentError?.(event.detail);
	element.addEventListener('document-change', changed);
	element.addEventListener('document-error', failed);
	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			element.locale = next.locale ?? 'en';
			element.readOnly = next.readOnly ?? false;
			element.theme = next.theme ?? 'auto';
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
