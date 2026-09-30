import type { DocumentModel } from '@christophervr/docx-core';
import type { StepBatch } from './collaboration';
import type { FileCommandDetail } from './file-commands';
import type { PageChangeDetail } from './page-sync';
import type { PresenceMessage } from './presence';
import type { RibbonAction } from './ribbon';

/** Every CustomEvent the `<docx-editor>` element (and its controllers) dispatches, with its detail. */
export interface DocxEditorEventMap {
	/** The document changed through editing, review actions or edits outside the main surface. */
	'document-change': CustomEvent<DocumentModel>;
	/** Loading, saving, printing or an edit failed. */
	'document-error': CustomEvent<Error>;
	/** Non-fatal notice, for example that a print layout could not be produced faithfully. */
	'document-warning': CustomEvent<string>;
	/** The user toggled read-only from the editor chrome. */
	'readonly-change': CustomEvent<boolean>;
	/** A ribbon control was activated. Bubbles from inside the shadow tree. */
	'ribbon-action': CustomEvent<RibbonAction>;
	/** The user changed which ribbon commands are shown (File > Customize Ribbon); persist the ids if you want them kept. */
	'ribbon-customize': CustomEvent<readonly string[]>;
	/** Cancelable: call `preventDefault()` to handle a File command (open/save/…) yourself. */
	'file-command': CustomEvent<FileCommandDetail>;
	/** Print Layout's current page or page count changed (approximate pagination, not Word's). */
	'page-change': CustomEvent<PageChangeDetail>;
	/** Unsaved-changes state flipped: true after an edit, false after save, load or `markClean()`. */
	'dirty-change': CustomEvent<boolean>;
	/** Local presence changed; transport it to other clients. */
	'presence-send': CustomEvent<PresenceMessage>;
	/** Local steps are pending; transport the batch to the collaboration authority. */
	'collaboration-send': CustomEvent<StepBatch>;
}

export type DocxEditorEventName = keyof DocxEditorEventMap;
/** Detail payload type of the named event. */
export type DocxEditorEventDetail<K extends DocxEditorEventName> =
	DocxEditorEventMap[K] extends CustomEvent<infer D> ? D : never;

/** Runtime list of every event name; the `satisfies` check keeps it a subset of the map. */
export const DOCX_EDITOR_EVENTS = [
	'document-change',
	'document-error',
	'document-warning',
	'readonly-change',
	'ribbon-action',
	'ribbon-customize',
	'file-command',
	'page-change',
	'dirty-change',
	'presence-send',
	'collaboration-send',
] as const satisfies readonly DocxEditorEventName[];

/** Typed listener for plain elements (such as the ribbon) that are not a `<docx-editor>`. */
export function on<K extends DocxEditorEventName>(
	target: EventTarget,
	name: K,
	listener: (event: DocxEditorEventMap[K]) => void,
): void {
	target.addEventListener(name, (event) => listener(event as DocxEditorEventMap[K]));
}

/**
 * Dispatches a typed, bubbling, composed CustomEvent. Returns false when a cancelable event was
 * canceled by a listener (same contract as `EventTarget.dispatchEvent`).
 */
export function emit<K extends DocxEditorEventName>(
	target: EventTarget,
	name: K,
	detail: DocxEditorEventDetail<K>,
	options: { cancelable?: boolean } = {},
): boolean {
	return target.dispatchEvent(
		new CustomEvent(name, {
			detail,
			bubbles: true,
			composed: true,
			...options,
		}),
	);
}
