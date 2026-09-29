export type ZoomFit = 'actual' | 'width' | 'page';

/**
 * The zoom percentage that makes a page fit its viewport: 100% for `actual`, the viewport's width
 * for `width`, and both dimensions for `page`. Sizes are in the same CSS-pixel units; the result
 * is a whole percent clamped to the 10-500% range Word allows.
 */
export function fitZoomPercent(
	mode: ZoomFit,
	viewport: { width: number; height: number },
	page: { width: number; height: number },
): number {
	if (mode === 'actual' || page.width <= 0 || page.height <= 0) return 100;
	const byWidth = viewport.width / page.width;
	const ratio = mode === 'width' ? byWidth : Math.min(byWidth, viewport.height / page.height);
	return Math.max(10, Math.min(500, Math.floor(ratio * 100)));
}
