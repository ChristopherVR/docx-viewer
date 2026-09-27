import type { LayoutColumns, LayoutPageGeometry } from './input.js';
import type { LayoutBlockBox, LayoutColumnBox, LayoutPageBox } from './result.js';

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
		return this.columnHeightPx - this.yPx;
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
	}
}
