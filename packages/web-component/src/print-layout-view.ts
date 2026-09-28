import type { DocumentModel } from '@christophervr/docx-core';
import {
	createCanvasMeasurer,
	layoutDocumentModel,
	type LayoutResult,
} from '@christophervr/docx-layout';
import { renderPrintLayout, type PictureUrl, type PrintLayoutHandle } from './print-layout.js';
import { decoratePages } from './print-header-footer';

const RELAYOUT_DEBOUNCE_MS = 150;

/**
 * Print Layout view: a read-only, paginated render kept in sync with the
 * editor model on a debounced relayout, rather than a second live editing
 * surface. Clicking a line asks the caller to move the real ProseMirror
 * cursor to the nearest editor position (see `print-layout.ts`'s
 * `resolveClick`), which is the chosen approach for keeping a single
 * source of truth for editing while still allowing click-to-place-cursor
 * from the paginated view.
 */
export interface PrintLayoutController {
	/** Mount this in the DOM alongside the continuous editing surface. */
	element: HTMLElement;
	setActive(active: boolean): void;
	scheduleRelayout(model: DocumentModel): void;
	/** Recomputes which page is most visible; call on the scroll container's `scroll` event. */
	refreshCurrentPage(): void;
	pageStatus(): { current: number; total: number } | null;
	/** Approximations collected by the last layout pass (see `LayoutResult.approximations`). */
	approximations(): string[];
	/** Lays out immediately (no debounce) and opens the browser print dialog for the paginated render. */
	print(model: DocumentModel, note: (message: string) => void): void;
	destroy(): void;
}

export function createPrintLayoutController(
	scrollContainer: HTMLElement,
	onRequestCursor: (blockId: string, offset: number) => void,
	pictureUrl?: PictureUrl,
): PrintLayoutController {
	let measurer = createCanvasMeasurer();
	let lastModel: DocumentModel | null = null;
	// Widths measured before a web font loads come from a fallback font; measure again once it has.
	const fonts = typeof document === 'undefined' ? undefined : document.fonts;
	const onFontsLoaded = () => {
		measurer = createCanvasMeasurer();
		if (lastModel && !element.hidden) relayout(lastModel);
	};
	fonts?.addEventListener?.('loadingdone', onFontsLoaded);
	const element = document.createElement('div');
	element.className = 'dve-print-pages';
	element.hidden = true;

	let handle: PrintLayoutHandle | null = null;
	let result: LayoutResult | null = null;
	let currentPage = 1;
	let timer: ReturnType<typeof setTimeout> | undefined;

	function relayout(model: DocumentModel) {
		lastModel = model;
		result = layoutDocumentModel(model, measurer);
		handle = renderPrintLayout(result, pictureUrl);
		decoratePages(
			model,
			result.pages,
			[...handle.element.querySelectorAll<HTMLElement>('.dve-print-page')],
			new Date(),
			pictureUrl,
		);
		element.replaceChildren(handle.element);
		refreshCurrentPage();
	}

	function onClick(event: MouseEvent) {
		if (!handle || !(event.target instanceof Element)) return;
		const hit = handle.resolveClick(event.target, event.clientX);
		if (hit) onRequestCursor(hit.blockId, hit.offset);
	}
	element.addEventListener('click', onClick);

	function refreshCurrentPage() {
		const pages = element.querySelectorAll<HTMLElement>('.dve-print-page');
		if (!pages.length) {
			currentPage = 1;
			return;
		}
		const anchor = scrollContainer.getBoundingClientRect().top + scrollContainer.clientHeight / 3;
		let closest = 0;
		let closestDistance = Number.POSITIVE_INFINITY;
		pages.forEach((page, index) => {
			const distance = Math.abs(page.getBoundingClientRect().top - anchor);
			if (distance < closestDistance) {
				closestDistance = distance;
				closest = index;
			}
		});
		currentPage = closest + 1;
	}

	return {
		element,
		setActive(active) {
			element.hidden = !active;
		},
		scheduleRelayout(model) {
			if (timer) clearTimeout(timer);
			timer = setTimeout(() => relayout(model), RELAYOUT_DEBOUNCE_MS);
		},
		refreshCurrentPage,
		pageStatus() {
			return result ? { current: currentPage, total: result.pages.length } : null;
		},
		approximations() {
			return result?.approximations ?? [];
		},
		print(model, note) {
			if (timer) clearTimeout(timer);
			relayout(model);
			printLayoutResult(result, note);
		},
		destroy() {
			if (timer) clearTimeout(timer);
			element.removeEventListener('click', onClick);
			fonts?.removeEventListener?.('loadingdone', onFontsLoaded);
			lastModel = null;
		},
	};
}

/** `@page`-sized print path: prints the current Print Layout render as-is. */
export function printLayoutResult(
	result: LayoutResult | null,
	note: (message: string) => void,
): void {
	if (!result || !result.pages.length) return;
	const first = result.pages[0];
	if (
		result.pages.some((page) => page.widthPx !== first.widthPx || page.heightPx !== first.heightPx)
	)
		note(
			'Printing uses the first page’s size for @page; mixed page sizes across sections print at one size.',
		);
	const style = document.createElement('style');
	style.textContent = `@page { size: ${first.widthPx / 96}in ${first.heightPx / 96}in; margin: 0; }`;
	document.head.append(style);
	window.print();
	style.remove();
}
