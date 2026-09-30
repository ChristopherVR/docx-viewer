import { PageCursor } from './page-cursor.js';
import { placeParagraph } from './flow-paragraph.js';
import { placeTable } from './flow-table.js';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout.js';
import { layoutRow, stackParagraphs } from './table-layout.js';
import { suppressesSpacing } from './keep-rules.js';
import type { TextMeasurer } from './measure.js';
import type { LayoutBlock, LayoutDocumentInput, LayoutSection } from './input.js';
import type { LayoutPageBox, LayoutResult } from './result.js';
import { lineBoxesFor, type Exclusion } from './wrap.js';
import { expectDefined } from './expect-defined.js';

/**
 * Section-break approximation: every section here starts a fresh page, even
 * one marked `continuous`. Word's continuous section breaks can change page
 * geometry/column count mid-page without a page break, which this engine
 * does not attempt to represent (see docs/parity-roadmap.md §3).
 */
const CONTINUOUS_BREAK_NOTE =
	'Continuous section breaks are rendered as page breaks; changing page size, margins or column count without starting a new page is not modeled.';

class SectionFlow {
	private readonly cursor: PageCursor;
	private readonly blocks: LayoutBlock[];
	private readonly measurer: TextMeasurer;
	private readonly note: (message: string) => void;
	private readonly paragraphCache = new Map<string, ParagraphLayoutResult>();
	private readonly firstRowHeightCache = new Map<string, number>();

	constructor(
		section: LayoutSection,
		pages: LayoutPageBox[],
		measurer: TextMeasurer,
		note: (m: string) => void,
		sectionIndex = 0,
		private readonly exclusions?: ReadonlyMap<number, Exclusion[]>,
	) {
		this.blocks = section.blocks;
		this.measurer = measurer;
		this.note = note;
		this.cursor = new PageCursor(
			pages,
			section.page,
			section.columns ?? { count: 1, gapPx: 0 },
			sectionIndex,
		);
		if (section.break === 'continuous') note(CONTINUOUS_BREAK_NOTE);
	}

	private paragraphLayout(index: number): ParagraphLayoutResult {
		const key = `${index}:${this.cursor.columnWidthPx}`;
		let cached = this.paragraphCache.get(key);
		if (!cached) {
			cached = layoutParagraph(
				this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>,
				this.cursor.columnWidthPx,
				this.measurer,
				this.note,
			);
			this.paragraphCache.set(key, cached);
		}
		return cached;
	}

	private firstRowHeight(index: number): number {
		const key = `${index}:${this.cursor.columnWidthPx}`;
		let cached = this.firstRowHeightCache.get(key);
		if (cached === undefined) {
			const table = this.blocks[index] as Extract<LayoutBlock, { kind: 'table' }>;
			const [firstRow] = table.rows;
			cached = firstRow
				? layoutRow(firstRow, this.cursor.columnWidthPx, this.measurer, this.note).heightPx
				: 0;
			this.firstRowHeightCache.set(key, cached);
		}
		return cached;
	}

	private spacingBeforeFor(index: number): number {
		const block = this.blocks[index];
		if (block?.kind !== 'paragraph') return 0;
		const layout = this.paragraphLayout(index);
		const previous = index > 0 ? this.blocks[index - 1] : undefined;
		if (previous?.kind === 'paragraph' && suppressesSpacing(previous, block)) return 0;
		return layout.spacingBeforePx;
	}

	/** Minimum height that must fit before `index` may start, walking a `keepNext` chain. */
	private requiredKeepHeight(index: number, depth = 0): number {
		if (index >= this.blocks.length || depth > 64) return 0;
		const block = expectDefined(this.blocks[index], 'keep-chain block index');
		if (block.kind === 'table') return this.firstRowHeight(index);
		const layout = this.paragraphLayout(index);
		const before = this.spacingBeforeFor(index);
		const own =
			before +
			(block.keepLines
				? layout.contentHeightPx + layout.spacingAfterPx
				: (layout.lines[0]?.heightPx ?? 0));
		return block.keepNext ? own + this.requiredKeepHeight(index + 1, depth + 1) : own;
	}

	/**
	 * The paragraph laid out around wrapped pictures on the current page, when any of them reaches
	 * the paragraph's lines; lines continuing onto a later page keep these widths.
	 */
	private wrappedLayout(index: number, spacingBeforePx: number): ParagraphLayoutResult | undefined {
		const page = this.cursor.page;
		const exclusions = this.exclusions?.get(page.index);
		if (!exclusions) return undefined;
		const plain = this.paragraphLayout(index);
		const topPx = page.marginTopPx + this.cursor.y + spacingBeforePx;
		const bottomPx = topPx + plain.contentHeightPx;
		if (!exclusions.some((item) => item.yPx < bottomPx && item.yPx + item.heightPx > topPx))
			return undefined;
		return layoutParagraph(
			this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>,
			this.cursor.columnWidthPx,
			this.measurer,
			this.note,
			lineBoxesFor(
				exclusions,
				topPx,
				page.marginLeftPx + this.cursor.column.xPx,
				this.cursor.columnWidthPx,
			),
		);
	}

	run(): void {
		for (let index = 0; index < this.blocks.length; index++) {
			const block = expectDefined(this.blocks[index], 'flow block index');
			if (block.kind === 'table') {
				placeTable(this.cursor, block, this.measurer, this.note);
				continue;
			}
			const spacingBeforePx = this.spacingBeforeFor(index);
			const layout = this.wrappedLayout(index, spacingBeforePx) ?? this.paragraphLayout(index);
			if (block.footnotes?.length)
				this.cursor.holdFootnotes(
					block.id,
					block.footnotes.map((note) => {
						const stacked = stackParagraphs(
							note.paragraphs,
							this.cursor.columnWidthPx,
							this.measurer,
							this.note,
						);
						return { id: note.id, yPx: 0, heightPx: stacked.heightPx, paragraphs: stacked.boxes };
					}),
				);
			const requiredTogetherPx = block.keepNext
				? this.requiredKeepHeight(index)
				: block.keepLines
					? spacingBeforePx + layout.contentHeightPx + layout.spacingAfterPx
					: 0;
			placeParagraph(
				this.cursor,
				{
					paragraph: block,
					layout,
					spacingBeforePx,
					requiredTogetherPx,
					reflow: (token) =>
						layoutParagraph(
							block,
							this.cursor.columnWidthPx,
							this.measurer,
							this.note,
							undefined,
							token,
						),
				},
				this.note,
			);
		}
	}
}

/**
 * Word starts an even-page (odd-page) section on the next even (odd) page, leaving a blank page
 * when needed. The blank page belongs to the previous section.
 */
function insertParityBlank(pages: LayoutPageBox[], section: LayoutSection): void {
	const previous = pages.at(-1);
	if (!previous || (section.break !== 'evenPage' && section.break !== 'oddPage')) return;
	const nextNumber = pages.length + 1;
	const wanted = section.break === 'evenPage' ? 0 : 1;
	if (nextNumber % 2 === wanted) return;
	pages.push({
		...previous,
		index: pages.length,
		pageInSection: previous.pageInSection + 1,
		columns: previous.columns.map((column) => ({ ...column, blocks: [] })),
	});
}

/** Shifts each page's content down for centered or bottom-aligned sections (`w:vAlign`). */
function alignVertically(
	pages: LayoutPageBox[],
	section: LayoutSection,
	note: (m: string) => void,
) {
	const align = section.verticalAlign;
	if (!align || align === 'top') return;
	if (align === 'both') {
		note('Vertically justified sections are laid out top-aligned.');
		return;
	}
	for (const page of pages) {
		const available = page.heightPx - page.marginTopPx - page.marginBottomPx;
		const used = Math.max(
			0,
			...page.columns.flatMap((column) => column.blocks.map((block) => block.yPx + block.heightPx)),
		);
		const shift = (available - used) / (align === 'center' ? 2 : 1);
		if (shift <= 0) continue;
		for (const column of page.columns) for (const block of column.blocks) block.yPx += shift;
	}
}

/** Paginates an already-adapted engine input. See `layout.ts` for the DocumentModel-facing entry point. */
export function layoutSections(
	input: LayoutDocumentInput,
	measurer: TextMeasurer,
	exclusions?: ReadonlyMap<number, Exclusion[]>,
): LayoutResult {
	const approximations = new Set<string>();
	const note = (message: string) => approximations.add(message);
	const pages: LayoutPageBox[] = [];
	input.sections.forEach((section, index) => {
		insertParityBlank(pages, section);
		const first = pages.length;
		new SectionFlow(section, pages, measurer, note, index, exclusions).run();
		alignVertically(pages.slice(first), section, note);
	});
	return { pages, approximations: [...approximations] };
}
