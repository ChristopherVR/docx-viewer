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

/**
 * Word's horizontal ruler above the page: inch ticks counted from the left margin, the margins
 * shaded, and markers for the current paragraph's first-line, left and right indents. This ruler
 * displays them; dragging a marker is not supported (use Layout > Indent or the Paragraph dialog).
 */
export function createRuler(): HTMLElement {
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
	const positions = markerPositions(geometry);
	part('marker-first-line').style.left = `${positions.firstLine}px`;
	part('marker-left').style.left = `${positions.left}px`;
	part('marker-right').style.left = `${positions.right}px`;
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
