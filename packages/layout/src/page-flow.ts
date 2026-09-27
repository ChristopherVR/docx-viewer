import { PageCursor } from './page-cursor.js';
import { placeParagraph } from './flow-paragraph.js';
import { placeTable } from './flow-table.js';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout.js';
import { layoutRow } from './table-layout.js';
import { suppressesSpacing } from './keep-rules.js';
import type { TextMeasurer } from './measure.js';
import type { LayoutBlock, LayoutDocumentInput, LayoutSection } from './input.js';
import type { LayoutPageBox, LayoutResult } from './result.js';

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
	private readonly paragraphCache = new Map<number, ParagraphLayoutResult>();
	private readonly firstRowHeightCache = new Map<number, number>();

	constructor(
		section: LayoutSection,
		pages: LayoutPageBox[],
		measurer: TextMeasurer,
		note: (m: string) => void,
	) {
		this.blocks = section.blocks;
		this.measurer = measurer;
		this.note = note;
		this.cursor = new PageCursor(pages, section.page, section.columns ?? { count: 1, gapPx: 0 });
		if (section.break === 'continuous') note(CONTINUOUS_BREAK_NOTE);
	}

	private paragraphLayout(index: number): ParagraphLayoutResult {
		let cached = this.paragraphCache.get(index);
		if (!cached) {
			cached = layoutParagraph(
				this.blocks[index] as Extract<LayoutBlock, { kind: 'paragraph' }>,
				this.cursor.columnWidthPx,
				this.measurer,
				this.note,
			);
			this.paragraphCache.set(index, cached);
		}
		return cached;
	}

	private firstRowHeight(index: number): number {
		let cached = this.firstRowHeightCache.get(index);
		if (cached === undefined) {
			const table = this.blocks[index] as Extract<LayoutBlock, { kind: 'table' }>;
			cached = table.rows.length
				? layoutRow(table.rows[0], this.cursor.columnWidthPx, this.measurer, this.note).heightPx
				: 0;
			this.firstRowHeightCache.set(index, cached);
		}
		return cached;
	}

	private spacingBeforeFor(index: number): number {
		const block = this.blocks[index];
		if (block.kind !== 'paragraph') return 0;
		const layout = this.paragraphLayout(index);
		const previous = index > 0 ? this.blocks[index - 1] : undefined;
		if (previous?.kind === 'paragraph' && suppressesSpacing(previous, block)) return 0;
		return layout.spacingBeforePx;
	}

	/** Minimum height that must fit before `index` may start, walking a `keepNext` chain. */
	private requiredKeepHeight(index: number, depth = 0): number {
		if (index >= this.blocks.length || depth > 64) return 0;
		const block = this.blocks[index];
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

	run(): void {
		for (let index = 0; index < this.blocks.length; index++) {
			const block = this.blocks[index];
			if (block.kind === 'table') {
				placeTable(this.cursor, block, this.measurer, this.note);
				continue;
			}
			const layout = this.paragraphLayout(index);
			const spacingBeforePx = this.spacingBeforeFor(index);
			const requiredTogetherPx = block.keepNext
				? this.requiredKeepHeight(index)
				: block.keepLines
					? spacingBeforePx + layout.contentHeightPx + layout.spacingAfterPx
					: 0;
			placeParagraph(
				this.cursor,
				{ paragraph: block, layout, spacingBeforePx, requiredTogetherPx },
				this.note,
			);
		}
	}
}

/** Paginates an already-adapted engine input. See `layout.ts` for the DocumentModel-facing entry point. */
export function layoutSections(input: LayoutDocumentInput, measurer: TextMeasurer): LayoutResult {
	const approximations = new Set<string>();
	const note = (message: string) => approximations.add(message);
	const pages: LayoutPageBox[] = [];
	for (const section of input.sections) new SectionFlow(section, pages, measurer, note).run();
	return { pages, approximations: [...approximations] };
}
