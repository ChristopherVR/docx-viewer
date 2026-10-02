import type { LayoutResult } from '@christophervr/ooxml-core/docx/layout';

/** Line numbering of a section, as the section model gives it. */
export interface PrintLineNumbering {
	countBy: number;
	start: number;
	restart: 'newPage' | 'newSection' | 'continuous';
	distanceTwips?: number;
}

/** Advances the body-line counter across columns/pages, skipping suppressed paragraphs. */
export function lineNumberLabels(
	page: LayoutResult['pages'][number],
	settings: PrintLineNumbering,
	counter: { next: number; section: number },
	suppressed?: ReadonlySet<string>,
): HTMLElement[] {
	const newSection = counter.section !== page.sectionIndex;
	if (
		counter.section === -1 ||
		settings.restart === 'newPage' ||
		(newSection && settings.restart === 'newSection')
	)
		counter.next = settings.start;
	counter.section = page.sectionIndex;
	const labels: HTMLElement[] = [];
	const gap = (settings.distanceTwips ?? 360) / 15;
	for (const column of page.columns)
		for (const block of column.blocks) {
			if (block.kind !== 'paragraph' || suppressed?.has(block.blockId)) continue;
			for (const line of block.lines) {
				const number = counter.next++;
				if (number % settings.countBy !== 0) continue;
				const label = document.createElement('span');
				label.className = 'dve-print-line-number';
				label.textContent = String(number);
				Object.assign(label.style, {
					left: `${page.marginLeftPx + column.xPx - gap - 40}px`,
					top: `${page.marginTopPx + block.yPx + line.yPx}px`,
					height: `${line.heightPx}px`,
					lineHeight: `${line.heightPx}px`,
				});
				labels.push(label);
			}
		}
	return labels;
}
