import type { LayoutTabStop } from './input.js';
import { DEFAULT_TAB_STOP_PX } from './units.js';

export interface TabPlacement {
	widthPx: number;
	leader?: LayoutTabStop['leader'];
}

/**
 * Where a tab at `positionPx` (from the paragraph's text margin) advances to, as Word does:
 * the next custom stop after it, else the hanging indent (an implicit stop for list labels),
 * else the next default stop. `followingPx` is the width of the text after the tab up to the
 * next tab or line end, used by right, center and decimal stops; `beforeDecimalPx` is the part
 * of it before the first decimal separator.
 */
export function placeTab(
	positionPx: number,
	stops: readonly LayoutTabStop[],
	hangingStopPx: number | undefined,
	followingPx: number,
	beforeDecimalPx: number,
): TabPlacement {
	const epsilon = 0.5;
	const custom = stops
		.filter((stop) => stop.align !== 'clear' && stop.align !== 'bar')
		.sort((a, b) => a.posPx - b.posPx)
		.find((stop) => stop.posPx > positionPx + epsilon);
	if (custom) {
		const distance = custom.posPx - positionPx;
		const shift =
			custom.align === 'right' || custom.align === 'end'
				? followingPx
				: custom.align === 'center'
					? followingPx / 2
					: custom.align === 'decimal'
						? beforeDecimalPx
						: 0;
		return {
			widthPx: Math.max(0, distance - shift),
			...(custom.leader && custom.leader !== 'none' ? { leader: custom.leader } : {}),
		};
	}
	if (hangingStopPx !== undefined && hangingStopPx > positionPx + epsilon)
		return { widthPx: hangingStopPx - positionPx };
	const next = (Math.floor(positionPx / DEFAULT_TAB_STOP_PX) + 1) * DEFAULT_TAB_STOP_PX;
	return { widthPx: next - positionPx };
}
