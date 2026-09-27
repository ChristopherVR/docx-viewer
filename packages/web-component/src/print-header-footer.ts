import {
	dateFieldResult,
	fieldName,
	formatNoteNumber,
	type Block,
	type DocumentModel,
	type HeaderFooterContent,
	type HeaderFooterSlots,
	type Paragraph,
	type TextRun,
} from '@christophervr/docx-core';
import type { LayoutPageBox } from '@christophervr/docx-layout';

/** Page facts a header or footer field can show. */
export interface PageFieldValues {
	page: string;
	numPages: string;
	sectionPages: string;
	/** When the layout was produced; DATE and TIME fields update to it, as Word updates them. */
	now?: Date;
}

/** Word page numbers per laid-out page, honoring each section's restart value and number format. */
export function pageNumbers(model: DocumentModel, pages: LayoutPageBox[]): string[] {
	let previous = 0;
	return pages.map((page) => {
		const numbering = model.sections?.[page.sectionIndex]?.pageNumbering;
		const value =
			page.pageInSection === 0 && numbering?.start !== undefined ? numbering.start : previous + 1;
		previous = value;
		return formatNoteNumber(value, numbering?.format ?? 'decimal');
	});
}

/**
 * The header or footer Word shows on a page: the first-page slot on a section's first page when
 * `titlePg` is set, the even slot on even pages when `evenAndOddHeaders` is on, else the default.
 * Sections without their own reference inherit the previous section's.
 */
export function headerFooterForPage(
	model: DocumentModel,
	page: LayoutPageBox,
	pageNumber: number,
	kind: 'headers' | 'footers',
): HeaderFooterContent | undefined {
	const sections = model.sections ?? [];
	const section = sections[page.sectionIndex];
	if (!section) return undefined;
	const slot: keyof HeaderFooterSlots =
		page.pageInSection === 0 && section.titlePage
			? 'first'
			: model.evenAndOddHeaders && pageNumber % 2 === 0
				? 'even'
				: 'default';
	for (let index = page.sectionIndex; index >= 0; index--) {
		const content = sections[index]?.[kind]?.[slot];
		if (content) return content;
	}
	return undefined;
}

/** Recalculates page fields; other fields keep the result Word last saved. */
export function fieldDisplayText(run: TextRun, values: PageFieldValues): string {
	if (!run.field) return run.text;
	const name = fieldName(run.field.instr);
	if (name === 'PAGE') return values.page;
	if (name === 'NUMPAGES') return values.numPages;
	if (name === 'SECTIONPAGES') return values.sectionPages;
	if (name === 'DATE' || name === 'TIME')
		return dateFieldResult(name, run.field.instr, values.now ?? new Date());
	return run.text;
}

function runElement(run: TextRun, values: PageFieldValues): HTMLElement {
	const span = document.createElement('span');
	span.textContent = fieldDisplayText(run, values);
	if (run.field) span.dataset.field = fieldName(run.field.instr);
	if (run.bold) span.style.fontWeight = '700';
	if (run.italic) span.style.fontStyle = 'italic';
	if (run.underline) span.style.textDecoration = 'underline';
	if (run.fontSize) span.style.fontSize = `${run.fontSize}pt`;
	if (run.fontFamily) span.style.fontFamily = `"${run.fontFamily.replace(/["\\]/g, '')}"`;
	if (run.color && /^#[0-9a-f]{6}$/i.test(run.color)) span.style.color = run.color;
	return span;
}

function paragraphElement(paragraph: Paragraph, values: PageFieldValues): HTMLElement {
	const element = document.createElement('p');
	if (paragraph.align) element.style.textAlign = paragraph.align;
	if (paragraph.direction) element.dir = paragraph.direction;
	for (const run of paragraph.runs) if (!run.break) element.append(runElement(run, values));
	return element;
}

function blockElement(block: Block, values: PageFieldValues): HTMLElement {
	if (block.type === 'paragraph') return paragraphElement(block, values);
	const table = document.createElement('table');
	for (const row of block.rows) {
		const tr = document.createElement('tr');
		for (const cell of row) {
			const td = document.createElement('td');
			if (cell.gridSpan && cell.gridSpan > 1) td.colSpan = cell.gridSpan;
			for (const paragraph of cell.paragraphs) td.append(paragraphElement(paragraph, values));
			tr.append(td);
		}
		table.append(tr);
	}
	return table;
}

/** Renders read-only header/footer content with page fields filled in for one page. */
export function renderHeaderFooter(
	content: HeaderFooterContent,
	values: PageFieldValues,
	kind: 'header' | 'footer',
): HTMLElement {
	const element = document.createElement('div');
	element.className = `dve-print-${kind}`;
	element.setAttribute('aria-label', kind === 'header' ? 'Header' : 'Footer');
	for (const block of content.blocks) element.append(blockElement(block, values));
	return element;
}

/** Adds each page's header and footer, positioned at the section's header/footer distances. */
export function decoratePages(
	model: DocumentModel,
	pages: LayoutPageBox[],
	sheets: HTMLElement[],
	now: Date = new Date(),
) {
	const numbers = pageNumbers(model, pages);
	const sectionPageCounts = new Map<number, number>();
	for (const page of pages)
		sectionPageCounts.set(page.sectionIndex, (sectionPageCounts.get(page.sectionIndex) ?? 0) + 1);
	pages.forEach((page, index) => {
		const sheet = sheets[index];
		if (!sheet) return;
		const values: PageFieldValues = {
			page: numbers[index],
			numPages: String(pages.length),
			sectionPages: String(sectionPageCounts.get(page.sectionIndex) ?? 1),
			now,
		};
		const section = model.sections?.[page.sectionIndex];
		const pageNumber = Number.parseInt(numbers[index], 10) || index + 1;
		for (const kind of ['header', 'footer'] as const) {
			const content = headerFooterForPage(model, page, pageNumber, `${kind}s`);
			if (!content) continue;
			const element = renderHeaderFooter(content, values, kind);
			const distanceTwips =
				kind === 'header' ? section?.headerDistanceTwips : section?.footerDistanceTwips;
			const distancePx = (distanceTwips ?? 720) / 15;
			element.style.left = `${page.marginLeftPx}px`;
			element.style.right = `${page.marginRightPx}px`;
			element.style[kind === 'header' ? 'top' : 'bottom'] = `${distancePx}px`;
			sheet.append(element);
		}
	});
}
