import type { LayoutBlockBox, LayoutLine, LayoutResult } from '@christophervr/docx-layout';

/** One clickable line, recorded for best-effort click-to-cursor mapping. */
interface LineHitBox {
	element: HTMLElement;
	blockId: string;
	sourceStart: number;
	sourceEnd: number;
}

export interface PrintLayoutHandle {
	/** The scrollable element holding one sheet per page; mount this in the DOM. */
	element: HTMLElement;
	pageCount: number;
	/**
	 * Best-effort click-to-cursor mapping: given the line element under a
	 * click and its horizontal offset, estimates a character position by
	 * distributing the line's source character range proportionally across
	 * its width. This is an approximation (it does not measure individual
	 * fragments), matching the honesty requirement that Print Layout is a
	 * read-only render synced from the editor, not a second editing surface.
	 */
	resolveClick(target: Element, clientX: number): { blockId: string; offset: number } | null;
}

function styleFragment(el: HTMLSpanElement, fragment: LayoutLine['fragments'][number]) {
	el.style.position = 'absolute';
	el.style.left = `${fragment.xPx}px`;
	el.style.whiteSpace = 'pre';
	if (fragment.bold) el.style.fontWeight = '700';
	if (fragment.italic) el.style.fontStyle = 'italic';
	if (fragment.fontFamily) el.style.fontFamily = fragment.fontFamily;
	if (fragment.fontSizePt) el.style.fontSize = `${fragment.fontSizePt}pt`;
	el.textContent = fragment.text;
}

function renderLine(line: LayoutLine, blockId: string, hitboxes: LineHitBox[]): HTMLElement {
	const lineEl = document.createElement('div');
	lineEl.className = 'dve-print-line';
	lineEl.style.top = `${line.yPx}px`;
	lineEl.style.height = `${line.heightPx}px`;
	for (const fragment of line.fragments) {
		if (!fragment.text && fragment.widthPx === 0) continue;
		const span = document.createElement('span');
		styleFragment(span, fragment);
		lineEl.append(span);
	}
	hitboxes.push({
		element: lineEl,
		blockId,
		sourceStart: line.sourceStart,
		sourceEnd: line.sourceEnd,
	});
	return lineEl;
}

function renderBlock(box: LayoutBlockBox, hitboxes: LineHitBox[]): HTMLElement {
	if (box.kind === 'paragraph') {
		const el = document.createElement('div');
		el.className = 'dve-print-block';
		el.style.top = `${box.yPx}px`;
		el.style.height = `${box.heightPx}px`;
		for (const line of box.lines) el.append(renderLine(line, box.blockId, hitboxes));
		return el;
	}
	const table = document.createElement('div');
	table.className = 'dve-print-block dve-print-table';
	table.style.top = `${box.yPx}px`;
	table.style.height = `${box.heightPx}px`;
	for (const row of box.rows) {
		const rowEl = document.createElement('div');
		rowEl.className = 'dve-print-row';
		rowEl.style.top = `${row.yPx}px`;
		rowEl.style.height = `${row.heightPx}px`;
		if (row.repeated) rowEl.dataset.repeated = 'true';
		// Column widths are not yet threaded through the model (see
		// docs/parity-roadmap.md §3); every cell renders as an equal flex share.
		for (const cell of row.cells) {
			const cellEl = document.createElement('div');
			cellEl.className = 'dve-print-cell';
			for (const paragraph of cell) cellEl.append(renderBlock(paragraph, hitboxes));
			rowEl.append(cellEl);
		}
		table.append(rowEl);
	}
	return table;
}

/** Pure, framework-neutral renderer: turns a `LayoutResult` into a DOM tree of page sheets. */
export function renderPrintLayout(result: LayoutResult): PrintLayoutHandle {
	const container = document.createElement('div');
	container.className = 'dve-print-pages';
	const hitboxes: LineHitBox[] = [];
	result.pages.forEach((page, pageIndex) => {
		const sheet = document.createElement('div');
		sheet.className = 'dve-print-page';
		sheet.dataset.pageIndex = String(pageIndex);
		sheet.style.width = `${page.widthPx}px`;
		sheet.style.height = `${page.heightPx}px`;
		for (const column of page.columns) {
			const columnEl = document.createElement('div');
			columnEl.className = 'dve-print-column';
			columnEl.style.top = `${page.marginTopPx}px`;
			columnEl.style.left = `${page.marginLeftPx + column.xPx}px`;
			columnEl.style.width = `${column.widthPx}px`;
			columnEl.style.height = `${page.heightPx - page.marginTopPx - page.marginBottomPx}px`;
			for (const block of column.blocks) columnEl.append(renderBlock(block, hitboxes));
			sheet.append(columnEl);
		}
		container.append(sheet);
	});
	return {
		element: container,
		pageCount: result.pages.length,
		resolveClick(target, clientX) {
			const hit = hitboxes.find((box) => box.element === target || box.element.contains(target));
			if (!hit) return null;
			const rect = hit.element.getBoundingClientRect();
			const ratio =
				rect.width > 0 ? Math.min(1, Math.max(0, (clientX - rect.left) / rect.width)) : 0;
			const span = Math.max(0, hit.sourceEnd - hit.sourceStart);
			const offset = hit.sourceStart + Math.round(span * ratio);
			return { blockId: hit.blockId, offset };
		},
	};
}
