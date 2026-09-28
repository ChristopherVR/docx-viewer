import {
	registerDocxEditor,
	type DocxEditorElement,
	type DocxEditorEventDetail,
	type DocxEditorEventName,
	type EditorLocaleInput,
	type EditorThemeMode,
	type PageChangeDetail,
	normalizeRibbonActions,
	type RibbonActionInput,
} from '@christophervr/docx-web-component';
import type { DocumentModel } from '@christophervr/docx-core';

/** Editor state a framework passes down as props. */
export interface EditorProps {
	// Framework props are `T | undefined` when unset, so `undefined` means "not provided".
	documentModel?: DocumentModel | undefined;
	readOnly?: boolean | undefined;
	/** Interface language: `en`, `fr`, `de`, `es`, `zh-CN`, or a tag such as `de-DE` that maps to one. */
	locale?: EditorLocaleInput | undefined;
	/** `light`, `dark`, or `auto` (default) to follow the OS color scheme. */
	theme?: EditorThemeMode | undefined;
	/** Left rail of page thumbnails (needs Print Layout). Default false. */
	showThumbnails?: boolean | undefined;
	/** Ribbon visibility. Default true. */
	showToolbar?: boolean | undefined;
	/** Ribbon controls to hide, by stable id (`RIBBON_ACTION_IDS`). Old English labels still work but warn. */
	hiddenActions?: readonly RibbonActionInput[] | undefined;
}
/** Editor callbacks; each is the framework-neutral form of one entry in `EDITOR_EVENT_NAMES`. */
export interface EditorEventOptions {
	onDocumentChange?: ((model: DocumentModel) => void) | undefined;
	onDocumentError?: ((error: Error) => void) | undefined;
	onPageChange?: ((detail: PageChangeDetail) => void) | undefined;
	onDirtyChange?: ((dirty: boolean) => void) | undefined;
}
export interface EditorOptions extends EditorProps, EditorEventOptions {}

/** Single source of truth for the prop keys every adapter forwards. */
export const EDITOR_PROP_KEYS = [
	'documentModel',
	'readOnly',
	'locale',
	'theme',
	'showThumbnails',
	'showToolbar',
	'hiddenActions',
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
	'page-change',
	'dirty-change',
] as const satisfies readonly DocxEditorEventName[];
export type EditorEventName = (typeof EDITOR_EVENT_NAMES)[number];
/** One handler per bound event; a missing key is a compile error in every adapter. */
export type EditorEventHandlers = {
	[K in EditorEventName]: ((detail: DocxEditorEventDetail<K>) => void) | undefined;
};

function sameList(a: readonly string[], b: readonly string[]): boolean {
	return a.length === b.length && a.every((item, index) => item === b[index]);
}
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
		onPageChange: handlers['page-change'],
		onDirtyChange: handlers['dirty-change'],
	};
}

export interface EditorHandle {
	readonly element: DocxEditorElement;
	load(input: Uint8Array | ArrayBuffer): Promise<void>;
	/** The saved document as a Blob. Does not clear `dirty`; call `markClean()` after persisting it. */
	save(): Promise<Blob>;
	/** Saves and downloads in the browser, then marks the document clean. */
	download(fileName?: string): Promise<void>;
	markClean(): void;
	readonly dirty: boolean;
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
	const paged = (event: CustomEvent<PageChangeDetail>) => options.onPageChange?.(event.detail);
	const dirtied = (event: CustomEvent<boolean>) => options.onDirtyChange?.(event.detail);
	element.addEventListener('document-change', changed);
	element.addEventListener('document-error', failed);
	element.addEventListener('page-change', paged);
	element.addEventListener('dirty-change', dirtied);
	const binding: EditorBinding = {
		element,
		update(next) {
			if (destroyed) return;
			options = next;
			element.locale = next.locale ?? 'en';
			element.readOnly = next.readOnly ?? false;
			element.theme = next.theme ?? 'auto';
			element.showThumbnails = next.showThumbnails ?? false;
			element.showToolbar = next.showToolbar ?? true;
			if (!sameList(element.hiddenActions, normalizeRibbonActions(next.hiddenActions ?? []).ids))
				element.hiddenActions = next.hiddenActions ?? [];
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
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		get dirty() {
			return element.dirty;
		},
		destroy() {
			if (destroyed) return;
			destroyed = true;
			element.removeEventListener('document-change', changed);
			element.removeEventListener('document-error', failed);
			element.removeEventListener('page-change', paged);
			element.removeEventListener('dirty-change', dirtied);
			element.remove();
		},
	};
	binding.update(initial);
	host.append(element);
	return binding;
}
