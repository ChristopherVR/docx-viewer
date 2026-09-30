import type { DocumentModel, SectionProperties } from '@christophervr/docx-core';
import { applyPageSetup, type PageSetupValues } from './page-setup-model';
import { emit } from './events';
import type { EditorHost } from './editor-host';
import { findLocalizedControl } from './localization';
import type { PrintLayoutController } from './print-layout-view';
import { applyPageStyles } from './ribbon-commands';
import { currentSectionIndex, insertSectionBreak, sectionsOf } from './section-commands';
import { newHeaderFooterId, withBlankHeaderFooter, withPageNumber } from './header-footer-commands';
import { pageSetupChange } from './page-setup-change';
import { pageSizeOf } from './page-size';
import type { RibbonAction } from './ribbon';
import { sectionLayoutJson } from './section-layout';
import type { StatusBar } from './status-bar';
import { fitZoomPercent, type ZoomFit } from './zoom-fit';

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
		if (!section) return;
		const next = pageSetupChange(model, index, section, key, value);
		this.dispatchSections(next);
	}

	/** Shows the selection's section settings in the Layout tab. */
	syncControls(): void {
		const view = this.host.view();
		const toolbar = this.host.toolbar();
		if (!view || !toolbar) return;
		const model = this.host.model();
		const section = sectionsOf(model)[currentSectionIndex(view, model)];
		if (!section) return;
		const setSelect = (label: string, value: string) => {
			const select = findLocalizedControl<HTMLSelectElement>(toolbar, label);
			if (select) select.value = value;
		};
		setSelect('Orientation', section.orientation);
		setSelect('Hyphenation', model.autoHyphenation ? 'auto' : 'none');
		setSelect('Page size', pageSizeOf(section) ?? '');
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

	/** The section holding the selection, for the Page Setup dialog. */
	currentSection(): SectionProperties | undefined {
		const view = this.host.view();
		const model = this.host.model();
		return view ? sectionsOf(model)[currentSectionIndex(view, model)] : undefined;
	}

	/** Applies the Page Setup dialog to the current section as one undoable step. */
	applyPageSetupValues(values: PageSetupValues): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const model = this.host.model();
		this.dispatchSections(applyPageSetup(model, currentSectionIndex(view, model), values));
	}

	/** Insert > Page Number: a PAGE field in the header or footer, creating the part when needed. */
	insertPageNumber(position: 'top' | 'bottom', align: 'left' | 'center' | 'right'): boolean {
		return this.changeHeaderFooter((model) =>
			withPageNumber(model, position, align, newHeaderFooterId),
		);
	}

	/** Insert > Header / Footer: an empty part when the document has none. */
	insertHeaderFooter(kind: 'header' | 'footer'): boolean {
		return this.changeHeaderFooter((model) =>
			withBlankHeaderFooter(model, kind === 'header' ? 'headers' : 'footers', newHeaderFooterId),
		);
	}

	private changeHeaderFooter(change: (model: DocumentModel) => DocumentModel): boolean {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return false;
		const next = change(this.host.model());
		// Header and footer content lives in the model, outside the editor document, so like in-place
		// header edits this is not part of Ctrl+Z history. The layout is recorded without a history step
		// so the document always has a `sections` attribute to rebuild the model from.
		this.host.setModel(next);
		this.dispatchSections(next, false);
		this.host.edited();
		return true;
	}

	/** Layout > Page Color (document-wide, undoable): hex with or without `#`, or `none` to clear. */
	setPageColor(value: string): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const hex = value.replace(/^#/, '').toUpperCase();
		view.dispatch(
			view.state.tr.setDocAttribute('pageColor', /^[0-9A-F]{6}$/.test(hex) ? hex : null),
		);
	}

	/** Layout > Hyphenation: `auto` sets `w:autoHyphenation`, `none` clears it (document-wide, undoable). */
	setHyphenation(value: string): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		view.dispatch(view.state.tr.setDocAttribute('autoHyphenation', value === 'auto'));
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
	private dispatchSections(next: DocumentModel, undoable = true): void {
		const view = this.host.view();
		if (!view) return;
		const { page } = next;
		view.dispatch(
			view.state.tr
				.setMeta('addToHistory', undoable)
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
		const columns = sections.length === 1 ? sections[0]?.columns : undefined;
		const multiple = columns && columns.count > 1;
		paper.style.columnCount = multiple ? String(columns.count) : '';
		paper.style.columnGap = multiple ? `${((columns.spacingTwips ?? 720) / 15) * this.zoom}px` : '';
	}

	/** Zoom to 100%, the page width or a whole page, from the section holding the selection. */
	zoomTo(mode: ZoomFit): void {
		if (mode === 'pages' && this.viewMode !== 'print') this.setViewMode('print');
		const canvas = this.host.paper()?.parentElement;
		const view = this.host.view();
		const model = this.host.model();
		const section = view ? sectionsOf(model)[currentSectionIndex(view, model)] : undefined;
		if (!canvas || !section) return this.setZoom(100);
		const style = getComputedStyle(canvas);
		const pad = (a: string, b: string) => (parseFloat(a) || 0) + (parseFloat(b) || 0);
		this.setZoom(
			fitZoomPercent(
				mode,
				{
					width: canvas.clientWidth - pad(style.paddingLeft, style.paddingRight),
					height: canvas.clientHeight - pad(style.paddingTop, style.paddingBottom),
				},
				{ width: section.pageWidthTwips / 15, height: section.pageHeightTwips / 15 },
			),
			mode === 'pages',
		);
	}

	/** `sideBySide` lays Print Layout's pages out two across (View > Multiple Pages). */
	setZoom(percent: number, sideBySide = false): void {
		this.zoom = percent / 100;
		this.host.paper()?.parentElement?.toggleAttribute('data-multipage', sideBySide);
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
			?.print(this.host.model(), (message) => emit(this.host.element, 'document-warning', message));
	}
}
