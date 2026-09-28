import type { DocumentModel } from '@christophervr/docx-core';
import type { EditorHost } from './editor-host';
import { findLocalizedControl } from './localization';
import type { PrintLayoutController } from './print-layout-view';
import { applyPageStyles } from './ribbon-commands';
import {
	currentSectionIndex,
	insertSectionBreak,
	sectionsOf,
	setColumns,
	setMargins,
	setOrientation,
	setPageNumbering,
	setTitlePage,
	setVerticalAlign,
} from './section-commands';
import type { RibbonAction } from './ribbon';
import { sectionLayoutJson } from './section-layout';
import type { StatusBar } from './status-bar';

export interface PageControllerHost extends EditorHost {
	paper(): HTMLElement | undefined;
	toolbar(): HTMLElement | undefined;
	statusBar(): StatusBar | undefined;
	printLayout(): PrintLayoutController | undefined;
	refreshControls(): void;
}

/** Page setup, section breaks, zoom, Web/Print Layout and printing. */
export class PageController {
	zoom = 1;
	viewMode: 'draft' | 'print' = 'draft';

	constructor(private readonly host: PageControllerHost) {}

	/** Page setup applies to the section holding the selection, as in Word, and is undoable. */
	changePageSetup(key: Extract<RibbonAction, { type: 'page' }>['key'], value: string): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const model = this.host.model();
		const index = currentSectionIndex(view, model);
		const section = sectionsOf(model)[index];
		const next =
			key === 'margin'
				? setMargins(model, index, value)
				: key === 'orientation'
					? setOrientation(model, index, value === 'landscape' ? 'landscape' : 'portrait')
					: key === 'columns'
						? setColumns(model, index, Math.max(1, Number(value) || 1))
						: key === 'numberFormat'
							? setPageNumbering(model, index, { format: value })
							: key === 'numberStart'
								? setPageNumbering(model, index, { restart: value === 'restart' })
								: key === 'verticalAlign'
									? setVerticalAlign(
											model,
											index,
											(['top', 'center', 'both', 'bottom'] as const).find(
												(item) => item === value,
											) ?? 'top',
										)
									: setTitlePage(model, index, !section.titlePage);
		this.dispatchSections(next);
	}

	/** Shows the selection's section settings in the Layout tab. */
	syncControls(): void {
		const view = this.host.view();
		const toolbar = this.host.toolbar();
		if (!view || !toolbar) return;
		const model = this.host.model();
		const section = sectionsOf(model)[currentSectionIndex(view, model)];
		const setSelect = (label: string, value: string) => {
			const select = findLocalizedControl<HTMLSelectElement>(toolbar, label);
			if (select) select.value = value;
		};
		setSelect('Orientation', section.orientation);
		setSelect('Vertical alignment', section.verticalAlign ?? 'top');
		findLocalizedControl<HTMLButtonElement>(toolbar, 'Different odd and even pages')?.setAttribute(
			'aria-pressed',
			String(Boolean(model.evenAndOddHeaders)),
		);
		setSelect('Columns', String(section.columns.count));
		setSelect('Page number format', section.pageNumbering?.format ?? 'decimal');
		setSelect(
			'Page numbering',
			section.pageNumbering?.start !== undefined ? 'restart' : 'continue',
		);
		findLocalizedControl<HTMLButtonElement>(toolbar, 'Different first page')?.setAttribute(
			'aria-pressed',
			String(Boolean(section.titlePage)),
		);
	}

	/** Header & Footer > Different Odd & Even Pages (document-wide, undoable). */
	toggleEvenOddHeaders(): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		view.dispatch(
			view.state.tr.setDocAttribute('evenAndOddHeaders', !view.state.doc.attrs.evenAndOddHeaders),
		);
	}

	insertSectionBreak(kind: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		try {
			this.dispatchSections(insertSectionBreak(view, this.host.model(), kind));
		} catch (cause) {
			this.host.reportError(cause);
		}
	}

	/** Records page geometry and section layout on the editor document as one undoable step. */
	private dispatchSections(next: DocumentModel): void {
		const view = this.host.view();
		if (!view) return;
		const { page } = next;
		view.dispatch(
			view.state.tr
				.setDocAttribute('pageWidth', page.width)
				.setDocAttribute('pageHeight', page.height)
				.setDocAttribute('marginTop', page.marginTop)
				.setDocAttribute('marginRight', page.marginRight)
				.setDocAttribute('marginBottom', page.marginBottom)
				.setDocAttribute('marginLeft', page.marginLeft)
				.setDocAttribute('sections', next.sections ? sectionLayoutJson(next.sections) : null),
		);
	}

	/** Page size and margins from the model; one multi-column section also shows its columns. */
	refreshPageStyles(): void {
		const paper = this.host.paper();
		if (!paper) return;
		const model = this.host.model();
		applyPageStyles(paper, model, this.zoom);
		const sections = model.sections ?? [];
		const columns = sections.length === 1 ? sections[0].columns : undefined;
		const multiple = columns && columns.count > 1;
		paper.style.columnCount = multiple ? String(columns.count) : '';
		paper.style.columnGap = multiple ? `${((columns.spacingTwips ?? 720) / 15) * this.zoom}px` : '';
	}

	setZoom(percent: number): void {
		this.zoom = percent / 100;
		this.refreshPageStyles();
		const toolbar = this.host.toolbar();
		const select = toolbar && findLocalizedControl<HTMLSelectElement>(toolbar, 'Zoom');
		if (select && [...select.options].some((option) => option.value === String(percent)))
			select.value = String(percent);
		this.host.statusBar()?.setZoom(percent);
	}

	/** Switches between the continuous editing surface and the paginated Print Layout render. */
	setViewMode(mode: 'draft' | 'print'): void {
		this.viewMode = mode;
		const printLayout = this.host.printLayout();
		printLayout?.setActive(mode === 'print');
		const paper = this.host.paper();
		if (paper) paper.hidden = mode === 'print';
		// Print Layout shows notes on their pages; the notes editing panel belongs to the editing view.
		paper?.parentElement?.toggleAttribute('data-print-view', mode === 'print');
		const toolbar = this.host.toolbar();
		const select = toolbar && findLocalizedControl<HTMLSelectElement>(toolbar, 'Layout view');
		if (select) select.value = mode;
		if (mode === 'print') printLayout?.scheduleRelayout(this.host.model());
		this.host.statusBar()?.setViewMode(mode);
		this.host.refreshControls();
	}

	/** Re-lays out Print Layout after a model change (debounced by the layout view). */
	relayout(): void {
		if (this.viewMode === 'print') this.host.printLayout()?.scheduleRelayout(this.host.model());
	}

	print(): void {
		this.host
			.printLayout()
			?.print(this.host.model(), (message) =>
				this.host.element.dispatchEvent(
					new CustomEvent('document-warning', { detail: message, bubbles: true, composed: true }),
				),
			);
	}
}
