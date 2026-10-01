import type { PageBorders } from '@christophervr/docx-core';
import { cssBorderSide } from './table-render';

/** The page geometry the border box is measured from (pixels at 96 DPI). */
export interface BorderPage {
	widthPx: number;
	heightPx: number;
	marginTopPx: number;
	marginRightPx: number;
	marginBottomPx: number;
	marginLeftPx: number;
	pageInSection: number;
}

const POINT = 4 / 3;

/**
 * A positioned box drawing a section's page borders on one printed sheet, or null when the borders
 * do not show on this page (`w:display`) or no side has a line. Each side sits `w:space` points from
 * the page edge (default) or from the text margins (`offsetFrom="text"`). Art borders are drawn as
 * plain lines, as the document model reports. It never takes pointer events.
 */
export function pageBorderBox(page: BorderPage, borders: PageBorders): HTMLElement | null {
	if (borders.display === 'firstPage' && page.pageInSection !== 0) return null;
	if (borders.display === 'notFirstPage' && page.pageInSection === 0) return null;
	const css = {
		top: cssBorderSide(borders.top),
		left: cssBorderSide(borders.left),
		bottom: cssBorderSide(borders.bottom),
		right: cssBorderSide(borders.right),
	};
	if (!Object.values(css).some(Boolean)) return null;
	const space = (side: 'top' | 'left' | 'bottom' | 'right') =>
		(borders[side]?.spacePoints ?? 24) * POINT;
	const fromText = borders.offsetFrom === 'text';
	const left = fromText ? page.marginLeftPx - space('left') : space('left');
	const top = fromText ? page.marginTopPx - space('top') : space('top');
	const right = fromText ? page.marginRightPx - space('right') : space('right');
	const bottom = fromText ? page.marginBottomPx - space('bottom') : space('bottom');
	const box = document.createElement('div');
	box.className = 'dve-print-page-border';
	box.setAttribute('aria-hidden', 'true');
	Object.assign(box.style, {
		position: 'absolute',
		boxSizing: 'border-box',
		pointerEvents: 'none',
		left: `${Math.max(0, left)}px`,
		top: `${Math.max(0, top)}px`,
		width: `${Math.max(0, page.widthPx - Math.max(0, left) - Math.max(0, right))}px`,
		height: `${Math.max(0, page.heightPx - Math.max(0, top) - Math.max(0, bottom))}px`,
		zIndex: borders.zOrder === 'back' ? '0' : '2',
	});
	for (const [side, value] of Object.entries(css))
		if (value) box.style.setProperty(`border-${side}`, value);
	return box;
}
