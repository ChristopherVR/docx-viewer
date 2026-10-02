import type { PageBorders } from '@christophervr/docx-core';
import {
	cssFontStack,
	ligatureCss,
	DEFAULT_FONT_SIZE_PT,
	type LayoutBlockBox,
	type LayoutLine,
	type LayoutParagraphFrame,
	type LayoutResult,
} from '@christophervr/ooxml-core/docx/layout';
import { renderTable } from './print-table';
import { lineNumberLabels, type PrintLineNumbering } from './print-line-numbers';
import { pageBorderBox } from './print-page-borders';
export type { PrintLineNumbering } from './print-line-numbers';

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

/** Resolves a package picture part to a displayable URL; undefined draws a placeholder box. */
export type PictureUrl = (partName: string, contentType: string) => string | undefined;

function pictureElement(
	object: { partName: string; contentType: string; widthPx: number; heightPx: number },
	pictureUrl: PictureUrl | undefined,
): HTMLElement {
	const url = object.partName ? pictureUrl?.(object.partName, object.contentType) : undefined;
	const element = document.createElement(url ? 'img' : 'div');
	if (url) {
		(element as HTMLImageElement).src = url;
		(element as HTMLImageElement).alt = '';
	} else element.classList.add('dve-print-picture-missing');
	element.classList.add('dve-print-picture');
	element.style.position = 'absolute';
	element.style.width = `${object.widthPx}px`;
	element.style.height = `${object.heightPx}px`;
	return element;
}

const LEADER_CHARACTERS: Record<string, string> = { dot: '.', hyphen: '-', middleDot: '·' };

/** A tab's leader: repeated characters (dots, hyphens) or a rule (underscore, heavy) across its width. */
function leaderElement(fragment: LayoutLine['fragments'][number]): HTMLElement {
	const el = document.createElement('span');
	styleFragment(el, { ...fragment, text: '' });
	el.classList.add('dve-print-leader');
	const scale = (fragment.textScalePercent ?? 100) / 100;
	el.style.width = `${scale ? fragment.widthPx / scale : fragment.widthPx}px`;
	el.style.overflow = 'hidden';
	el.style.textAlign = 'right';
	const character = LEADER_CHARACTERS[fragment.leader!];
	if (character) {
		// Enough characters to fill the tab; the overflow is clipped on the left.
		const sizePx = ((fragment.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * 96) / 72;
		el.textContent = character.repeat(
			Math.ceil(fragment.widthPx / ((scale || 1) * sizePx * 0.25)) + 1,
		);
		el.style.direction = 'rtl';
	} else {
		el.style.height = '1em';
		el.style.borderBottom = `${fragment.leader === 'heavy' ? 2 : 1}px solid currentColor`;
	}
	return el;
}

function styleFragment(el: HTMLSpanElement, fragment: LayoutLine['fragments'][number]) {
	el.style.position = 'absolute';
	el.style.left = `${fragment.xPx}px`;
	el.style.whiteSpace = 'pre';
	if (fragment.bold) el.style.fontWeight = '700';
	if (fragment.italic) el.style.fontStyle = 'italic';
	// The same metric-compatible stack the measurer used, so rendered text matches its line breaks.
	el.style.fontFamily = cssFontStack(fragment.fontFamily);
	el.style.fontSize = `${(fragment.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * (fragment.script ? 0.65 : 1)}pt`;
	const scale = (fragment.textScalePercent ?? 100) / 100;
	if (scale !== 1) {
		el.style.transform = `scaleX(${scale})`;
		el.style.transformOrigin = 'left center';
	}
	if (fragment.characterSpacingPx !== undefined)
		el.style.letterSpacing = `${scale ? fragment.characterSpacingPx / scale : 0}px`;
	if (fragment.kerningThresholdPt !== undefined)
		el.style.fontKerning =
			fragment.kerningThresholdPt > 0 &&
			(fragment.fontSizePt ?? DEFAULT_FONT_SIZE_PT) >= fragment.kerningThresholdPt
				? 'normal'
				: 'none';
	// The layout places each fragment's box so its baseline sits on the line's baseline.
	if (fragment.topPx !== undefined) el.style.top = `${fragment.topPx}px`;
	if (fragment.boxHeightPx !== undefined) el.style.lineHeight = `${fragment.boxHeightPx}px`;
	if (fragment.color) el.style.color = fragment.color;
	if (fragment.ligatures !== undefined) el.style.fontVariantLigatures = ligatureCss(fragment.ligatures);
	const lines = [fragment.underline && 'underline', fragment.strike && 'line-through'].filter(
		Boolean,
	);
	if (lines.length) el.style.textDecorationLine = lines.join(' ');
	el.textContent = fragment.text;
}

function renderLine(
	line: LayoutLine,
	blockId: string,
	hitboxes: LineHitBox[],
	pictureUrl: PictureUrl | undefined,
): HTMLElement {
	const lineEl = document.createElement('div');
	lineEl.className = 'dve-print-line';
	lineEl.style.top = `${line.yPx}px`;
	lineEl.style.height = `${line.heightPx}px`;
	for (const fragment of line.fragments) {
		if (fragment.object) {
			const picture = pictureElement(fragment.object, pictureUrl);
			picture.style.left = `${fragment.xPx}px`;
			picture.style.top = `${fragment.topPx ?? 0}px`;
			lineEl.append(picture);
			continue;
		}
		if (fragment.leader) {
			lineEl.append(leaderElement(fragment));
			continue;
		}
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

/** A paragraph's borders and shading, drawn behind its text across the box's height. */
function paragraphFrame(frame: LayoutParagraphFrame, heightPx: number): HTMLElement {
	const el = document.createElement('div');
	el.className = 'dve-print-paragraph-frame';
	Object.assign(el.style, {
		left: `${frame.leftPx}px`,
		width: `${frame.widthPx}px`,
		height: `${heightPx}px`,
	});
	if (frame.shading) el.style.background = frame.shading;
	for (const side of ['top', 'right', 'bottom', 'left'] as const) {
		const line = frame.borders?.[side];
		if (line)
			el.style.setProperty(`border-${side}`, `${line.widthPx}px ${line.style} ${line.color}`);
	}
	return el;
}

function renderBlock(
	box: LayoutBlockBox,
	hitboxes: LineHitBox[],
	pictureUrl: PictureUrl | undefined,
	columnWidthPx: number,
): HTMLElement {
	if (box.kind === 'paragraph') {
		const el = document.createElement('div');
		el.className = 'dve-print-block';
		el.style.top = `${box.yPx}px`;
		el.style.height = `${box.heightPx}px`;
		if (box.frame) el.append(paragraphFrame(box.frame, box.heightPx));
		for (const line of box.lines) el.append(renderLine(line, box.blockId, hitboxes, pictureUrl));
		return el;
	}
	return renderTable(box, columnWidthPx, (paragraph) =>
		renderBlock(paragraph, hitboxes, pictureUrl, columnWidthPx),
	);
}

/** The page's footnotes above its bottom margin, under Word's short separator rule. */
function footnoteArea(
	page: LayoutResult['pages'][number],
	hitboxes: LineHitBox[],
	pictureUrl: PictureUrl | undefined,
): HTMLElement {
	const notes = page.footnotes ?? [];
	const heightPx = notes.reduce((sum, note) => sum + note.heightPx, 0);
	const widthPx = page.widthPx - page.marginLeftPx - page.marginRightPx;
	const area = document.createElement('div');
	area.className = 'dve-print-footnotes';
	area.setAttribute('aria-label', 'Footnotes');
	Object.assign(area.style, {
		left: `${page.marginLeftPx}px`,
		width: `${widthPx}px`,
		top: `${page.heightPx - page.marginBottomPx - heightPx}px`,
		height: `${heightPx}px`,
	});
	const separator = document.createElement('div');
	separator.className = 'dve-print-footnote-separator';
	area.append(separator);
	for (const note of notes) {
		const noteEl = document.createElement('div');
		noteEl.className = 'dve-print-footnote';
		noteEl.dataset.noteId = note.id;
		Object.assign(noteEl.style, { top: `${note.yPx}px`, height: `${note.heightPx}px` });
		for (const paragraph of note.paragraphs)
			noteEl.append(renderBlock(paragraph, hitboxes, pictureUrl, widthPx));
		area.append(noteEl);
	}
	return area;
}

export interface PrintLayoutOptions {
	/** Line numbering by section index; sections without an entry print no numbers. */
	lineNumbers?: ReadonlyArray<PrintLineNumbering | undefined>;
	/** Page borders by section index; drawn on each page they apply to. */
	pageBorders?: ReadonlyArray<PageBorders | undefined>;
	/** Paragraphs with effective `w:suppressLineNumbers`; omitted from the counter. */
	suppressedLineNumberParagraphs?: ReadonlySet<string>;
}

/** Pure, framework-neutral renderer: turns a `LayoutResult` into a DOM tree of page sheets. */
export function renderPrintLayout(
	result: LayoutResult,
	pictureUrl?: PictureUrl,
	options: PrintLayoutOptions = {},
): PrintLayoutHandle {
	const container = document.createElement('div');
	container.className = 'dve-print-pages';
	const hitboxes: LineHitBox[] = [];
	const counter = { next: 1, section: -1 };
	result.pages.forEach((page, pageIndex) => {
		const sheet = document.createElement('div');
		sheet.className = 'dve-print-page';
		sheet.dataset.pageIndex = String(pageIndex);
		sheet.style.width = `${page.widthPx}px`;
		sheet.style.height = `${page.heightPx}px`;
		const borders = options.pageBorders?.[page.sectionIndex];
		const borderBox = borders ? pageBorderBox(page, borders) : null;
		if (borderBox) sheet.append(borderBox);
		for (const column of page.columns) {
			const previous = page.columns[page.columns.indexOf(column) - 1];
			if (page.columnSeparator && previous) {
				const rule = document.createElement('div');
				rule.className = 'dve-print-column-rule';
				Object.assign(rule.style, {
					position: 'absolute',
					left: `${page.marginLeftPx + (previous.xPx + previous.widthPx + column.xPx) / 2}px`,
					top: `${page.marginTopPx}px`,
					height: `${page.heightPx - page.marginTopPx - page.marginBottomPx}px`,
					borderLeft: '1px solid currentColor',
				});
				sheet.append(rule);
			}
			const columnEl = document.createElement('div');
			columnEl.className = 'dve-print-column';
			columnEl.style.top = `${page.marginTopPx}px`;
			columnEl.style.left = `${page.marginLeftPx + column.xPx}px`;
			columnEl.style.width = `${column.widthPx}px`;
			columnEl.style.height = `${page.heightPx - page.marginTopPx - page.marginBottomPx}px`;
			for (const block of column.blocks)
				columnEl.append(renderBlock(block, hitboxes, pictureUrl, column.widthPx));
			sheet.append(columnEl);
		}
		const numbering = options.lineNumbers?.[page.sectionIndex];
		if (numbering)
			sheet.append(
				...lineNumberLabels(page, numbering, counter, options.suppressedLineNumberParagraphs),
			);
		if (page.footnotes?.length) sheet.append(footnoteArea(page, hitboxes, pictureUrl));
		for (const float of page.floats ?? []) {
			const picture = pictureElement(float, pictureUrl);
			picture.classList.add('dve-print-float');
			if (float.behindText) picture.classList.add('dve-print-float-behind');
			picture.style.left = `${float.xPx}px`;
			picture.style.top = `${float.yPx}px`;
			sheet.append(picture);
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
