import type { EditorCore } from './editor-core';
import { emit } from './events';

/** Payload of the `page-change` event. Pages are 1-based and come from the approximate layout. */
export interface PageChangeDetail {
	page: number;
	pageCount: number;
}

/** Emits `page-change` only when the page or the page count really changed. */
export class PageTracker {
	private last?: PageChangeDetail;

	constructor(private readonly target: EventTarget) {}

	report(status: { current: number; total: number } | null): void {
		if (!status) return;
		if (this.last?.page === status.current && this.last.pageCount === status.total) return;
		this.last = { page: status.current, pageCount: status.total };
		emit(this.target, 'page-change', { ...this.last });
	}
}

/**
 * Single source for page state: Print Layout's `pageStatus()` feeds the navigator rail, the
 * `page-change` event and (through `refreshControls`) the status bar. Outside Print Layout there is
 * no pagination, so nothing is reported and the rail explains why.
 */
export function syncPageState(core: EditorCore): void {
	const { printLayout, navigator } = core.shell;
	const active = core.pages.viewMode === 'print';
	const status = active ? (printLayout?.pageStatus() ?? null) : null;
	navigator?.sync({ active, status, locale: core.locale });
	core.pageTracker.report(status);
}
