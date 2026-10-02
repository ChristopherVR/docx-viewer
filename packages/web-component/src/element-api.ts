import { saveDocx } from 'docx-core';
import type { EditorCore } from './editor-core';
import { reflectAttribute } from './editor-attributes';
import { downloadBytes, wordBlob } from './file-commands';
import {
	normalizeRibbonActions,
	type RibbonActionId,
	type RibbonActionInput,
} from './ribbon-action-ids';
import { emit } from './events';
import { applyViewOptions } from './view-options';

const HTMLElementBase: typeof HTMLElement =
	typeof HTMLElement === 'undefined' ? (class {} as typeof HTMLElement) : HTMLElement;

/**
 * Save, dirty-tracking and UI-customisation surface of `<docx-editor>`. It only reads and writes
 * the shared `EditorCore`, so the element class stays a thin lifecycle facade.
 */
export abstract class DocxEditorApi extends HTMLElementBase {
	protected abstract readonly core: EditorCore;
	abstract fileName: string;

	/** Shows the left rail of page thumbnails. Needs Print Layout; reflected to `show-thumbnails`. */
	get showThumbnails(): boolean {
		return this.core.viewOptions.showThumbnails;
	}
	set showThumbnails(value: boolean) {
		const next = Boolean(value);
		if (next !== this.core.viewOptions.showThumbnails) {
			this.core.viewOptions.showThumbnails = next;
			applyViewOptions(this.core);
		}
		reflectAttribute(this, 'show-thumbnails', this.showThumbnails);
	}

	/** Hides the ribbon (the title bar and status bar stay). Default true; `show-toolbar="false"`. */
	get showToolbar(): boolean {
		return this.core.viewOptions.showToolbar;
	}
	set showToolbar(value: boolean) {
		const next = Boolean(value);
		if (next !== this.core.viewOptions.showToolbar) {
			this.core.viewOptions.showToolbar = next;
			applyViewOptions(this.core);
		}
		reflectAttribute(this, 'show-toolbar', this.showToolbar);
	}

	/**
	 * Ribbon controls to hide, by stable id (`RIBBON_ACTION_IDS`, e.g. `'bold'`). Unknown ids are
	 * ignored. The old English labels (`'Bold'`) are still accepted for one release: they are mapped
	 * to ids and reported through a `document-warning` event. Reading returns ids only.
	 */
	get hiddenActions(): readonly RibbonActionId[] {
		return this.core.viewOptions.hiddenActions;
	}
	set hiddenActions(value: readonly RibbonActionInput[]) {
		const { ids: next, legacy } = normalizeRibbonActions(Array.isArray(value) ? value : []);
		if (legacy.length > 0)
			emit(
				this,
				'document-warning',
				`hiddenActions: English labels (${legacy.map((label) => `'${label}'`).join(', ')}) are deprecated and will be removed in the next release; use the stable ids from RIBBON_ACTION_IDS.`,
			);
		const current = this.core.viewOptions.hiddenActions;
		if (next.length === current.length && next.every((id, index) => id === current[index])) return;
		this.core.viewOptions.hiddenActions = next;
		applyViewOptions(this.core);
	}

	/** True after an edit until the document is saved through File > Save, `download()`, loaded or `markClean()`. */
	get dirty(): boolean {
		return this.core.dirtyState.dirty;
	}

	/** Declares the current content saved, for hosts that persist the result of `save()` themselves. */
	markClean(): void {
		this.core.dirtyState.set(false);
		this.core.shell.chrome?.setSaveState('saved');
	}

	/** Serialized document bytes: the loaded package's own writer, else a fresh DOCX from the model. */
	async saveBytes(): Promise<Uint8Array> {
		const { core } = this;
		// Only pass staged pictures when there are any: legacy DOC sessions take the model alone.
		const media = core.inserts.pendingMedia.size ? core.inserts.pendingMedia : undefined;
		if (core.loaded)
			return media ? core.loaded.save(core.model, media) : core.loaded.save(core.model);
		return saveDocx(core.model, media);
	}

	/** The saved document as a Blob. Does not change `dirty`; call `markClean()` after persisting it. */
	async save(): Promise<Blob> {
		return wordBlob(await this.saveBytes(), this.fileName);
	}

	/** Saves and starts a browser download, then marks the document clean. */
	async download(fileName?: string): Promise<void> {
		const name = fileName || this.fileName;
		downloadBytes(await this.saveBytes(), /\.[^./\\]+$/.test(name) ? name : `${name}.docx`);
		this.markClean();
	}
}
