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
} from 'docx-core';
import {
	floatPosition,
	paragraphFloats,
	type LayoutPageBox,
} from 'ooxml-core/docx/layout';
import type { PictureUrl } from './print-layout';

/** Page facts a header or footer field can show. */
export interface PageFieldValues {
	page: string;
	numPages: string;
	sectionPages: string;
	/** When the layout was produced; DATE and TIME fields update to it, as Word updates them. */
	now?: Date;
	/** Resolves header/footer pictures to displayable URLs. */
	pictureUrl?: PictureUrl;
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

/** An inline header/footer picture, or a placeholder box when its bytes are unavailable. */
function inlinePicture(image: NonNullable<TextRun['image']>, values: PageFieldValues): HTMLElement {
	const url = image.partName ? values.pictureUrl?.(image.partName, image.contentType) : undefined;
	const element = document.createElement(url ? 'img' : 'span');
	if (url) {
		(element as HTMLImageElement).src = url;
		(element as HTMLImageElement).alt = image.altText ?? '';
	} else element.className = 'dve-print-picture-missing';
	element.style.display = 'inline-block';
	element.style.verticalAlign = 'bottom';
	element.style.width = `${image.widthPx}px`;
	element.style.height = `${image.heightPx}px`;
	return element;
}

function runElement(run: TextRun, values: PageFieldValues): HTMLElement {
	if (run.image) return inlinePicture(run.image, values);
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
	// Floating pictures are placed on the sheet by `decoratePages`, not in the text flow.
	for (const run of paragraph.runs)
		if (!run.break && !run.image?.anchored && !run.image?.watermark)
			element.append(runElement(run, values));
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
/** Header/footer floating pictures, positioned on the sheet from their anchor frames. */
function floatingPictures(
	content: HeaderFooterContent,
	page: LayoutPageBox,
	anchorTopPx: number,
	values: PageFieldValues,
): HTMLElement[] {
	const paragraphs = content.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
	const column = { xPx: 0, widthPx: page.widthPx - page.marginLeftPx - page.marginRightPx };
	return paragraphs.flatMap((paragraph) =>
		paragraphFloats(paragraph).map((float) => {
			const { xPx, yPx } = floatPosition(float, page, column, { topPx: anchorTopPx, heightPx: 0 });
			const url = values.pictureUrl?.(float.partName, float.contentType);
			const element = document.createElement(url ? 'img' : 'div');
			if (url) (element as HTMLImageElement).src = url;
			else element.classList.add('dve-print-picture-missing');
			element.classList.add('dve-print-picture', 'dve-print-float');
			if (float.behindText) element.classList.add('dve-print-float-behind');
			Object.assign(element.style, {
				position: 'absolute',
				left: `${xPx}px`,
				top: `${yPx}px`,
				width: `${float.widthPx}px`,
				height: `${float.heightPx}px`,
			});
			return element;
		}),
	);
}

/** The header's text watermark as a centered, rotated layer behind the page text, or null. */
export function watermarkElement(
	content: HeaderFooterContent,
	page: Pick<LayoutPageBox, 'widthPx' | 'heightPx'>,
): HTMLElement | null {
	const spec = content.blocks
		.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
		.find((run) => run.image?.watermark)?.image?.watermark;
	if (!spec?.text) return null;
	const chars = Math.max(1, [...spec.text].length);
	const span = Math.min(page.widthPx * (spec.layout === 'diagonal' ? 0.78 : 0.7), 560);
	const layer = document.createElement('div');
	layer.className = 'dve-print-watermark';
	layer.setAttribute('aria-hidden', 'true');
	layer.textContent = spec.text;
	Object.assign(layer.style, {
		position: 'absolute',
		left: '0',
		top: '0',
		width: `${page.widthPx}px`,
		height: `${page.heightPx}px`,
		display: 'flex',
		alignItems: 'center',
		justifyContent: 'center',
		pointerEvents: 'none',
		userSelect: 'none',
		whiteSpace: 'nowrap',
		overflow: 'hidden',
		fontSize: `${Math.max(12, Math.min(220, span / (chars * 0.6)))}px`,
		fontWeight: '700',
		color: /^#[0-9a-f]{6}$/i.test(spec.color) ? spec.color : '#c0c0c0',
		opacity: spec.semitransparent ? '0.5' : '1',
		transform: spec.layout === 'diagonal' ? 'rotate(-45deg)' : 'none',
		zIndex: '0',
		...(spec.fontFamily
			? { fontFamily: `"${spec.fontFamily.replace(/["\\]/g, '')}", sans-serif` }
			: {}),
	});
	return layer;
}

export function decoratePages(
	model: DocumentModel,
	pages: LayoutPageBox[],
	sheets: HTMLElement[],
	now: Date = new Date(),
	pictureUrl?: PictureUrl,
) {
	const numbers = pageNumbers(model, pages);
	const sectionPageCounts = new Map<number, number>();
	for (const page of pages)
		sectionPageCounts.set(page.sectionIndex, (sectionPageCounts.get(page.sectionIndex) ?? 0) + 1);
	pages.forEach((page, index) => {
		const sheet = sheets[index];
		if (!sheet) return;
		const label = numbers[index] ?? String(index + 1);
		const values: PageFieldValues = {
			page: label,
			numPages: String(pages.length),
			sectionPages: String(sectionPageCounts.get(page.sectionIndex) ?? 1),
			now,
			...(pictureUrl ? { pictureUrl } : {}),
		};
		const section = model.sections?.[page.sectionIndex];
		const pageNumber = Number.parseInt(label, 10) || index + 1;
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
			// Anchors in a footer sit roughly one line above its bottom distance.
			const anchorTop = kind === 'header' ? distancePx : page.heightPx - distancePx - 20;
			sheet.append(...floatingPictures(content, page, anchorTop, values));
			if (kind === 'header') {
				const mark = watermarkElement(content, page);
				// First in the sheet so the page text paints over it, as a Word watermark sits behind text.
				if (mark) sheet.prepend(mark);
			}
		}
	});
}
