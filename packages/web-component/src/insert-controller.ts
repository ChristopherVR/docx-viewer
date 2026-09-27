import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel, PendingMediaPart } from '@christophervr/docx-core';
import { applyCharacterStyle } from './character-style-picker';
import { createLinkDialog, type LinkDialog } from './link-dialog';
import { followLinkAt } from './link-commands';
import { insertPicture, PICTURE_TYPES, stagePicture } from './picture-commands';
import type { RibbonAction } from './ribbon';
import type { EditorLocale } from './localization';
import { focusView } from './focus-view';

export interface InsertControllerHost {
	view(): EditorView | undefined;
	model(): DocumentModel;
	/** The editable content width in CSS pixels, used to fit inserted pictures. */
	contentWidth(): number;
	paper(): HTMLElement | undefined;
	toolbar(): HTMLElement | undefined;
	reportError(error: Error): void;
}

/** Picture, link, character style and hidden-text commands shared by every editor instance. */
export class InsertController {
	readonly linkDialog: LinkDialog;
	readonly pictureInput: HTMLInputElement;
	/** Bytes for pictures inserted since the document was loaded; passed to save. */
	readonly pendingMedia = new Map<string, PendingMediaPart>();
	private showHidden = false;

	constructor(private readonly host: InsertControllerHost) {
		this.linkDialog = createLinkDialog({
			getView: () => host.view(),
			onError: (error) => host.reportError(error),
		});
		this.pictureInput = document.createElement('input');
		this.pictureInput.type = 'file';
		this.pictureInput.accept = Object.keys(PICTURE_TYPES).join(',');
		this.pictureInput.hidden = true;
		this.pictureInput.className = 'dve-picture-input';
		this.pictureInput.setAttribute('aria-label', 'Insert picture');
		this.pictureInput.addEventListener('change', () => void this.insertSelectedPicture());
	}

	/** Handles the actions this controller owns; returns false for everything else. */
	handle(action: RibbonAction): boolean {
		const view = this.host.view();
		const editable = Boolean(view?.editable);
		if (action.type === 'insertPicture') {
			if (editable) this.pictureInput.click();
		} else if (action.type === 'link') {
			if (editable) this.linkDialog.open();
		} else if (action.type === 'characterStyle') {
			if (view?.editable) applyCharacterStyle(view, action.value);
			focusView(view);
		} else if (action.type === 'showHidden') this.setShowHidden(!this.showHidden);
		else return false;
		return true;
	}

	/** Ctrl+K opens the link dialog; Ctrl/Cmd+Click follows a link like Word. */
	plugins(): Plugin[] {
		return [
			keymap({
				'Mod-k': (_state, _dispatch, view) => {
					if (view?.editable) this.linkDialog.open();
					return true;
				},
			}),
		];
	}

	handleClick(view: EditorView, pos: number, event: MouseEvent): boolean {
		return (event.ctrlKey || event.metaKey) && followLinkAt(view, pos);
	}

	/** Media bytes for rendering: pictures inserted in this session first, then the loaded package. */
	media(partName: string, loaded?: ReadonlyMap<string, Uint8Array>): Uint8Array | undefined {
		return this.pendingMedia.get(partName)?.bytes ?? loaded?.get(partName);
	}

	reset(): void {
		this.pendingMedia.clear();
		if (this.linkDialog.isOpen) this.linkDialog.close();
	}

	setLocale(locale: EditorLocale): void {
		this.linkDialog.setLocale(locale);
	}

	/** Re-applies the hidden-text display state after the paper is rebuilt. */
	syncPaper(): void {
		const paper = this.host.paper();
		if (paper) paper.toggleAttribute('data-show-hidden', this.showHidden);
		this.host
			.toolbar()
			?.querySelector('[aria-label="Show hidden text"], [data-localearialabel="Show hidden text"]')
			?.setAttribute('aria-pressed', String(this.showHidden));
	}

	private setShowHidden(show: boolean): void {
		this.showHidden = show;
		this.syncPaper();
	}

	private async insertSelectedPicture(): Promise<void> {
		const file = this.pictureInput.files?.[0];
		this.pictureInput.value = '';
		const view = this.host.view();
		if (!file || !view?.editable) return;
		try {
			const staged = await stagePicture(file, this.host.contentWidth());
			this.pendingMedia.set(staged.image.partName, staged.media);
			insertPicture(view, staged.image);
			focusView(view);
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
		}
	}
}
