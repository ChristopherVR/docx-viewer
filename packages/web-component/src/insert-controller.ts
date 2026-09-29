import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DocumentModel, PendingMediaPart } from '@christophervr/docx-core';
import { applyCharacterStyle } from './character-style-picker';
import { createLinkDialog, type LinkDialog } from './link-dialog';
import { createPictureDialog, type PictureDialog } from './picture-dialog';
import { NodeSelection } from 'prosemirror-state';
import { followLinkAt } from './link-commands';
import { insertPicture, PICTURE_TYPES, stagePicture } from './picture-commands';
import type { RibbonAction } from './ribbon';
import { translate, type EditorLocale } from './localization';
import { focusView } from './focus-view';
import { insertTableOfContents, updateTableOfContents } from './toc-commands';

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
	readonly pictureDialog: PictureDialog;
	readonly pictureInput: HTMLInputElement;
	/** Bytes for pictures inserted since the document was loaded; passed to save. */
	readonly pendingMedia = new Map<string, PendingMediaPart>();
	private showHidden = false;
	private showMarks = false;
	private showGridlines = false;

	constructor(private readonly host: InsertControllerHost) {
		this.linkDialog = createLinkDialog({
			getView: () => host.view(),
			onError: (error) => host.reportError(error),
		});
		this.pictureDialog = createPictureDialog(() => host.view());
		this.pictureInput = document.createElement('input');
		this.pictureInput.type = 'file';
		this.pictureInput.accept = [...Object.keys(PICTURE_TYPES), 'image/svg+xml'].join(',');
		this.pictureInput.hidden = true;
		this.pictureInput.className = 'dve-picture-input';
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
		} else if (action.type === 'formatPicture') {
			const selection = view?.state.selection;
			if (editable && selection instanceof NodeSelection && selection.node.type.name === 'image')
				this.pictureDialog.open(selection.from);
		} else if (action.type === 'characterStyle') {
			if (view?.editable) applyCharacterStyle(view, action.value);
			focusView(view);
		} else if (action.type === 'showHidden') this.setShowHidden(!this.showHidden);
		else if (action.type === 'gridlines') {
			this.showGridlines = !this.showGridlines;
			this.syncPaper();
		} else if (action.type === 'showMarks') {
			this.showMarks = !this.showMarks;
			this.syncPaper();
		} else if (action.type === 'toc') {
			if (!view || !editable) return true;
			try {
				if (action.key === 'insert') insertTableOfContents(view, this.host.model());
				else if (!updateTableOfContents(view, this.host.model()))
					this.host.reportError(new Error('This document has no table of contents to update.'));
			} catch (error) {
				this.host.reportError(error instanceof Error ? error : new Error(String(error)));
			}
			focusView(view);
		} else return false;
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
		if (this.pictureDialog.isOpen) this.pictureDialog.close();
	}

	setLocale(locale: EditorLocale): void {
		this.linkDialog.setLocale(locale);
		this.pictureDialog.setLocale(locale);
		this.pictureInput.setAttribute('aria-label', translate(locale, 'Insert picture'));
	}

	/** Re-applies the hidden-text display state after the paper is rebuilt. */
	syncPaper(): void {
		const paper = this.host.paper();
		if (paper) {
			paper.toggleAttribute('data-show-hidden', this.showHidden);
			paper.toggleAttribute('data-show-marks', this.showMarks);
			paper.toggleAttribute('data-show-gridlines', this.showGridlines);
		}
		this.host
			.toolbar()
			?.querySelector(
				'[aria-label="Show paragraph marks"], [data-localearialabel="Show paragraph marks"]',
			)
			?.setAttribute('aria-pressed', String(this.showMarks));
		this.host
			.toolbar()
			?.querySelector('[aria-label="Gridlines"], [data-localearialabel="Gridlines"]')
			?.setAttribute('aria-pressed', String(this.showGridlines));
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
			if (staged.svg && staged.image.svgPartName)
				this.pendingMedia.set(staged.image.svgPartName, staged.svg);
			insertPicture(view, staged.image);
			focusView(view);
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
		}
	}
}
