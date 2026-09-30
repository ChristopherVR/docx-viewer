import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { createFontDialog, type FormatDialog } from './font-dialog';
import type { EditorLocale } from './localization';
import { createBookmarkDialog } from './bookmark-dialog';
import { createBordersDialog } from './borders-dialog';
import { createCaptionDialog } from './caption-dialog';
import { createCrossReferenceDialog } from './cross-reference-dialog';
import { blockPageNumbers } from './toc-commands';
import { focusView } from './focus-view';
import { createPageSetupDialog, type PageSetupHost } from './page-setup-dialog';
import { createParagraphDialog } from './paragraph-dialog';
import { createZoomDialog, type ZoomHost } from './zoom-dialog';
import { createTablePropertiesDialog } from './table-properties-dialog';
import { createLineNumberDialog, type LineNumberHost } from './line-number-dialog';
import { createDropCapDialog } from './drop-cap-dialog';
import { createColumnsDialog, type ColumnsHost } from './columns-dialog';
import { createMultilevelListDialog } from './multilevel-list-dialog';

export type FormatDialogKind =
	| 'font'
	| 'paragraph'
	| 'bookmark'
	| 'pageSetup'
	| 'caption'
	| 'crossReference'
	| 'zoom'
	| 'borders'
	| 'tableProperties'
	| 'lineNumbers'
	| 'dropCap'
	| 'columns'
	| 'multilevelList';

export interface FormatDialogsHost {
	view(): EditorView | undefined;
	/** Header/footer edits use the main document's history; notes use their own. */
	historyView?(): EditorView | undefined;
	model(): DocumentModel;
	canDefineList?(): boolean;
	/** Page Setup edits the section holding the selection. */
	pageSetup?: Omit<PageSetupHost, 'restoreFocus'>;
	/** View > Zoom: the current zoom and how to change it. */
	zoom?: Omit<ZoomHost, 'restoreFocus'>;
	lineNumbers?: Omit<LineNumberHost, 'restoreFocus'>;
	columns?: Omit<ColumnsHost, 'restoreFocus'>;
}

/** The Font and Paragraph dialogs opened from the ribbon's group launchers. */
export class FormatDialogs {
	private readonly dialogs: Record<FormatDialogKind, FormatDialog>;

	constructor(host: FormatDialogsHost) {
		this.dialogs = {
			font: createFontDialog(
				() => host.view(),
				() => host.historyView?.() ?? host.view(),
			),
			paragraph: createParagraphDialog(
				() => host.view(),
				() => host.model(),
			),
			bookmark: createBookmarkDialog(() => host.view()),
			caption: createCaptionDialog(
				() => host.view(),
				() => host.model(),
			),
			crossReference: createCrossReferenceDialog(
				() => host.view(),
				() => host.model(),
				(id) => blockPageNumbers(host.model()).get(id),
			),
			zoom: createZoomDialog({
				...(host.zoom ?? { percent: () => 100, setPercent: () => {}, fit: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			pageSetup: createPageSetupDialog({
				...(host.pageSetup ?? { section: () => undefined, canEdit: () => false, apply: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			borders: createBordersDialog(() => host.view()),
			tableProperties: createTablePropertiesDialog(() => host.view()),
			lineNumbers: createLineNumberDialog({
				...(host.lineNumbers ?? {
					section: () => undefined,
					canEdit: () => false,
					apply: () => {},
				}),
				restoreFocus: () => focusView(host.view()),
			}),
			dropCap: createDropCapDialog(() => host.view()),
			columns: createColumnsDialog({
				...(host.columns ?? { section: () => undefined, canEdit: () => false, apply: () => {} }),
				restoreFocus: () => focusView(host.view()),
			}),
			multilevelList: createMultilevelListDialog(
				() => host.view(),
				() => host.model(),
				() => host.canDefineList?.() ?? true,
			),
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
