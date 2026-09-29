import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { createFontDialog, type FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import { createBookmarkDialog } from './bookmark-dialog';
import { createParagraphDialog } from './paragraph-dialog';

export type FormatDialogKind = 'font' | 'paragraph' | 'bookmark';

export interface FormatDialogsHost {
	view(): EditorView | undefined;
	model(): DocumentModel;
}

/** The Font and Paragraph dialogs opened from the ribbon's group launchers. */
export class FormatDialogs {
	private readonly dialogs: Record<FormatDialogKind, FormatDialog>;

	constructor(host: FormatDialogsHost) {
		this.dialogs = {
			font: createFontDialog(() => host.view()),
			paragraph: createParagraphDialog(
				() => host.view(),
				() => host.model(),
			),
			bookmark: createBookmarkDialog(() => host.view()),
		};
	}

	get elements(): HTMLElement[] {
		return Object.values(this.dialogs).map((dialog) => dialog.element);
	}

	get list(): FormatDialog[] {
		return Object.values(this.dialogs);
	}

	/** Opens `kind`, closing the other one first. */
	open(kind: FormatDialogKind): void {
		for (const [name, dialog] of Object.entries(this.dialogs))
			if (name !== kind && dialog.isOpen) dialog.close();
		this.dialogs[kind].open();
	}

	closeAll(): void {
		for (const dialog of this.list) if (dialog.isOpen) dialog.close();
	}

	setLocale(locale: EditorLocale): void {
		for (const dialog of this.list) dialog.setLocale(locale);
	}
}
