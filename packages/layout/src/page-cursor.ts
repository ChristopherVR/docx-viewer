import type { LayoutColumns, LayoutPageGeometry } from './input.js';
import type {
	LayoutBlockBox,
	LayoutColumnBox,
	LayoutFootnoteBox,
	LayoutPageBox,
} from './result.js';

/** Space for the short rule above a page's footnotes. */
export const FOOTNOTE_SEPARATOR_PX = 12;

/**
 * Mutable fill-position bookkeeping for one section: which page/column is
 * currently being filled and how much vertical space remains in it. A new
 * `PageCursor` is created per section (see `page-flow.ts`'s approximation
 * that a continuous section break still starts a fresh page).
 */
export class PageCursor {
	readonly pages: LayoutPageBox[];
	readonly columnWidthPx: number;
	readonly columnHeightPx: number;
	private geometry: LayoutPageGeometry;
	private columns: LayoutColumns;
	private columnIndex = 0;
	private yPx = 0;
	private readonly sectionIndex: number;
	private pageInSection = 0;
	/** Height taken by footnotes (and their separator) at the bottom of the current page. */
	private reservedPx = 0;
	/** Footnotes of the paragraph about to be placed; they join the page its first box lands on. */
	private pending: { blockId: string; notes: LayoutFootnoteBox[]; heightPx: number } | undefined;

	constructor(
		pages: LayoutPageBox[],
		geometry: LayoutPageGeometry,
		columns: LayoutColumns,
		sectionIndex = 0,
	) {
		this.sectionIndex = sectionIndex;
		this.pages = pages;
		this.geometry = geometry;
		this.columns = columns;
		const contentWidth = geometry.widthPx - geometry.marginLeftPx - geometry.marginRightPx;
		this.columnWidthPx = Math.max(
			1,
			(contentWidth - columns.gapPx * (columns.count - 1)) / columns.count,
		);
		this.columnHeightPx = Math.max(
			1,
			geometry.heightPx - geometry.marginTopPx - geometry.marginBottomPx,
		);
		this.pushPage();
	}

	private pushPage() {
		const columnBoxes: LayoutColumnBox[] = [];
		for (let i = 0; i < this.columns.count; i++)
			columnBoxes.push({
				xPx: i * (this.columnWidthPx + this.columns.gapPx),
				widthPx: this.columnWidthPx,
				blocks: [],
			});
		this.pages.push({
			index: this.pages.length,
			sectionIndex: this.sectionIndex,
			pageInSection: this.pageInSection++,
			widthPx: this.geometry.widthPx,
			heightPx: this.geometry.heightPx,
			marginTopPx: this.geometry.marginTopPx,
			marginRightPx: this.geometry.marginRightPx,
			marginBottomPx: this.geometry.marginBottomPx,
			marginLeftPx: this.geometry.marginLeftPx,
			columns: columnBoxes,
		});
		this.columnIndex = 0;
		this.yPx = 0;
		this.reservedPx = 0;
	}

	get page(): LayoutPageBox {
		return this.pages.at(-1)!;
	}
	get column(): LayoutColumnBox {
		return this.page.columns[this.columnIndex];
	}
	get y(): number {
		return this.yPx;
	}
	get atColumnTop(): boolean {
		return this.yPx === 0;
	}
	remainingHeightPx(): number {
		return this.columnHeightPx - this.yPx - this.reservedPx - this.pendingHeightPx();
	}

	private pendingHeightPx(): number {
		if (!this.pending) return 0;
		return this.pending.heightPx + (this.page.footnotes?.length ? 0 : FOOTNOTE_SEPARATOR_PX);
	}

	/**
	 * Reserves room for a paragraph's footnotes on whichever page its first box is placed, so the
	 * notes print on the same page as their reference, as Word does.
	 */
	holdFootnotes(blockId: string, notes: LayoutFootnoteBox[]): void {
		const heightPx = notes.reduce((sum, note) => sum + note.heightPx, 0);
		this.pending = notes.length ? { blockId, notes, heightPx } : undefined;
	}

	newPage(): void {
		this.pushPage();
	}
	/** Advances to the next column, or a new page when the section has no more columns. */
	newColumn(): void {
		if (this.columnIndex + 1 < this.columns.count) {
			this.columnIndex++;
			this.yPx = 0;
		} else {
			this.pushPage();
		}
	}

	/** Places a box at the current fill position (sets its `yPx`) and advances by `advancePx`. */
	place(box: LayoutBlockBox, advancePx: number): void {
		box.yPx = this.yPx;
		this.column.blocks.push(box);
		this.yPx += advancePx;
		if (this.pending && this.pending.blockId === box.blockId) {
			this.reservedPx += this.pendingHeightPx();
			const page = this.page;
			let top = page.footnotes?.reduce((sum, note) => sum + note.heightPx, 0) ?? 0;
			for (const note of this.pending.notes) {
				(page.footnotes ??= []).push({ ...note, yPx: top });
				top += note.heightPx;
			}
			this.pending = undefined;
		}
	}
}
