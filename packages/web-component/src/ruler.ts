import type { ParagraphFormat } from './paragraph-format';

/** Everything the ruler draws, in CSS pixels at 100% zoom (96 per inch). */
export interface RulerGeometry {
	pageWidth: number;
	marginLeft: number;
	marginRight: number;
	/** Current paragraph's indents, measured from the margins. */
	indentLeft: number;
	indentRight: number;
	/** Extra first-line offset from the left indent: positive first line, negative hanging. */
	firstLine: number;
}

export const PX_PER_INCH = 96;

/** Marker positions from the page's left edge: first-line, left indent, right indent. */
export function markerPositions(geometry: RulerGeometry): {
	firstLine: number;
	left: number;
	right: number;
} {
	const left = geometry.marginLeft + geometry.indentLeft;
	return {
		firstLine: left + geometry.firstLine,
		left,
		right: geometry.pageWidth - geometry.marginRight - geometry.indentRight,
	};
}

/** Whole-inch labels of the ruler: 0 at the left margin, counting toward the right margin. */
export function inchLabels(geometry: RulerGeometry): Array<{ at: number; text: string }> {
	const labels: Array<{ at: number; text: string }> = [];
	const width = geometry.pageWidth - geometry.marginLeft - geometry.marginRight;
	for (let inch = 1; inch * PX_PER_INCH < width; inch++)
		labels.push({ at: geometry.marginLeft + inch * PX_PER_INCH, text: String(inch) });
	return labels;
}

export type RulerMarker = 'first-line' | 'left' | 'right';

/** The paragraph fields a marker dragged to `x` (px from the page's left edge) changes. */
export function markerChange(
	marker: RulerMarker,
	x: number,
	geometry: RulerGeometry,
): Partial<Pick<ParagraphFormat, 'leftInches' | 'rightInches' | 'special' | 'specialInches'>> {
	const snap = (px: number) => Math.round(px / SNAP_PX) * SNAP_PX;
	const inches = (px: number) => Math.round((px / PX_PER_INCH) * 1000) / 1000;
	const positions = markerPositions(geometry);
	if (marker === 'right') {
		const edge = geometry.pageWidth - geometry.marginRight;
		const indent = Math.min(Math.max(snap(edge - x), 0), edge - positions.left - SNAP_PX);
		return { rightInches: inches(indent) };
	}
	if (marker === 'left') {
		// Like Word's left-indent marker, the first-line offset travels with it.
		const limit = geometry.pageWidth - geometry.marginRight - geometry.indentRight - SNAP_PX;
		const indent = Math.min(
			Math.max(snap(x - geometry.marginLeft), 0),
			limit - geometry.marginLeft,
		);
		return { leftInches: inches(indent) };
	}
	const offset = snap(x - positions.left);
	const first = Math.max(offset, -positions.left);
	return {
		special: first > 0 ? 'firstLine' : first < 0 ? 'hanging' : 'none',
		specialInches: inches(Math.abs(first)),
	};
}

const SNAP_PX = PX_PER_INCH / 16;
const geometries = new WeakMap<HTMLElement, RulerGeometry>();

/** Lets the three markers be dragged; `onChange` gets the result once, when the pointer is released. */
function enableDragging(
	ruler: HTMLElement,
	onChange: (change: ReturnType<typeof markerChange>) => void,
): void {
	for (const name of ['first-line', 'left', 'right'] as const) {
		const marker = ruler.querySelector<HTMLElement>(`.dve-ruler-marker-${name}`)!;
		marker.addEventListener('pointerdown', (event) => {
			const geometry = geometries.get(ruler);
			if (!geometry || event.button !== 0) return;
			event.preventDefault();
			marker.setPointerCapture?.(event.pointerId);
			const zoom = Number(ruler.style.getPropertyValue('--dve-zoom')) || 1;
			const at = (e: PointerEvent) => (e.clientX - ruler.getBoundingClientRect().left) / zoom;
			const move = (e: PointerEvent) => {
				const preview = markerChange(name, at(e), geometry);
				const shifted = { ...geometry };
				if (preview.leftInches !== undefined) shifted.indentLeft = preview.leftInches * PX_PER_INCH;
				if (preview.rightInches !== undefined)
					shifted.indentRight = preview.rightInches * PX_PER_INCH;
				if (preview.special)
					shifted.firstLine =
						(preview.special === 'hanging' ? -1 : 1) * (preview.specialInches ?? 0) * PX_PER_INCH;
				paintMarkers(ruler, shifted);
			};
			const finish = (e: PointerEvent) => {
				marker.removeEventListener('pointermove', move);
				marker.removeEventListener('pointerup', finish);
				marker.removeEventListener('pointercancel', finish);
				onChange(e.type === 'pointerup' ? markerChange(name, at(e), geometry) : {});
			};
			marker.addEventListener('pointermove', move);
			marker.addEventListener('pointerup', finish);
			marker.addEventListener('pointercancel', finish);
		});
	}
}

function paintMarkers(ruler: HTMLElement, geometry: RulerGeometry): void {
	const positions = markerPositions(geometry);
	const at = (name: string) => ruler.querySelector<HTMLElement>(`.dve-ruler-marker-${name}`)!;
	at('first-line').style.left = `${positions.firstLine}px`;
	at('left').style.left = `${positions.left}px`;
	at('right').style.left = `${positions.right}px`;
}

/**
 * Word's horizontal ruler above the page: inch ticks counted from the left margin, the margins
 * shaded, and draggable markers for the current paragraph's first-line, left and right indents.
 * Dragging snaps to sixteenths of an inch and applies as one undoable step on release. Margin
 * edges and tab stops are not draggable.
 */
export function createRuler(
	onChange?: (change: ReturnType<typeof markerChange>) => void,
): HTMLElement {
	const ruler = document.createElement('div');
	ruler.className = 'dve-ruler';
	ruler.setAttribute('role', 'img');
	ruler.setAttribute('aria-label', 'Ruler');
	for (const name of ['margin-left', 'margin-right', 'labels']) {
		const part = document.createElement('div');
		part.className = `dve-ruler-${name}`;
		ruler.append(part);
	}
	for (const name of ['first-line', 'left', 'right']) {
		const marker = document.createElement('div');
		marker.className = `dve-ruler-marker dve-ruler-marker-${name}`;
		ruler.append(marker);
	}
	if (onChange) enableDragging(ruler, onChange);
	return ruler;
}

/** Positions the ruler's parts for `geometry`; `zoom` matches the page's CSS zoom. */
export function updateRuler(ruler: HTMLElement, geometry: RulerGeometry, zoom: number): void {
	const part = (name: string) => ruler.querySelector<HTMLElement>(`.dve-ruler-${name}`)!;
	ruler.style.width = `${geometry.pageWidth}px`;
	ruler.style.setProperty('--dve-zoom', String(zoom));
	ruler.style.setProperty('--dve-ruler-origin', `${geometry.marginLeft}px`);
	const left = part('margin-left');
	left.style.width = `${geometry.marginLeft}px`;
	const right = part('margin-right');
	right.style.width = `${geometry.marginRight}px`;
	geometries.set(ruler, geometry);
	paintMarkers(ruler, geometry);
	const labels = part('labels');
	const wanted = inchLabels(geometry);
	if (labels.dataset.key !== JSON.stringify(wanted)) {
		labels.dataset.key = JSON.stringify(wanted);
		labels.replaceChildren(
			...wanted.map(({ at, text }) => {
				const label = document.createElement('span');
				label.textContent = text;
				label.style.left = `${at}px`;
				return label;
			}),
		);
	}
}
