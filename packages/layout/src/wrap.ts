import type { LineBoxFn } from './paragraph-layout.js';
import type { LayoutPageBox } from './result.js';

/** Word's default distance between wrapped text and a picture's sides (`distL`/`distR`, 1/8 in). */
const SIDE_DISTANCE_PX = 12;
/** Narrower than this, a side is left empty and the line moves below the picture. */
const MIN_SIDE_WIDTH_PX = 48;

/** Where text may not go on a page, in page coordinates (CSS pixels). */
export interface Exclusion {
	xPx: number;
	yPx: number;
	widthPx: number;
	heightPx: number;
	/** Top-and-bottom wrapping keeps both sides empty. */
	topAndBottom: boolean;
}

/** Areas that wrapped floating pictures take from the text, per page index. */
export function wrapExclusions(pages: LayoutPageBox[]): Map<number, Exclusion[]> {
	const exclusions = new Map<number, Exclusion[]>();
	for (const page of pages) {
		const list = (page.floats ?? [])
			.filter((float) => float.wrap && float.wrap !== 'none')
			.map((float) => ({
				xPx: float.xPx - SIDE_DISTANCE_PX,
				yPx: float.yPx,
				widthPx: float.widthPx + SIDE_DISTANCE_PX * 2,
				heightPx: float.heightPx,
				topAndBottom: float.wrap === 'topAndBottom',
			}));
		if (list.length) exclusions.set(page.index, list);
	}
	return exclusions;
}

/** Whether two exclusion maps place the same areas (to stop re-laying out once floats settle). */
export function sameExclusions(a: Map<number, Exclusion[]>, b: Map<number, Exclusion[]>): boolean {
	const key = (map: Map<number, Exclusion[]>) =>
		JSON.stringify(
			[...map].map(([page, list]) => [
				page,
				list.map((item) => [Math.round(item.xPx), Math.round(item.yPx), item.topAndBottom]),
			]),
		);
	return key(a) === key(b);
}

/**
 * Line boxes for a paragraph whose first line starts at `topPx` (page coordinates) in a column
 * spanning `columnLeftPx`..`columnLeftPx + columnWidthPx`: lines beside a picture use the wider
 * side of it (Word's "largest side"), and lines that cannot fit beside it move below it.
 */
export function lineBoxesFor(
	exclusions: readonly Exclusion[],
	topPx: number,
	columnLeftPx: number,
	columnWidthPx: number,
): LineBoxFn {
	return (yPx, heightPx) => {
		const top = topPx + yPx;
		let leftInsetPx = 0;
		let rightInsetPx = 0;
		let gapBeforePx = 0;
		for (const exclusion of exclusions) {
			const overlapsLine =
				exclusion.yPx < top + heightPx && exclusion.yPx + exclusion.heightPx > top;
			const start = exclusion.xPx - columnLeftPx;
			const end = start + exclusion.widthPx;
			if (!overlapsLine || end <= 0 || start >= columnWidthPx) continue;
			const left = Math.max(0, start);
			const right = Math.max(0, columnWidthPx - end);
			if (exclusion.topAndBottom || Math.max(left, right) < MIN_SIDE_WIDTH_PX) {
				gapBeforePx = Math.max(gapBeforePx, exclusion.yPx + exclusion.heightPx - top);
				continue;
			}
			if (left >= right) rightInsetPx = Math.max(rightInsetPx, columnWidthPx - start);
			else leftInsetPx = Math.max(leftInsetPx, end);
		}
		return { leftInsetPx, rightInsetPx, gapBeforePx };
	};
}
