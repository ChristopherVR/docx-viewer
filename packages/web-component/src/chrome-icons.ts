/** Small inline SVG icons for the window chrome, drawn with `currentColor` so they follow the theme. */
const paths = {
	save: 'M4 3h9l3 3v11H4z M7 3v4h6V3 M7 17v-5h6v5',
	undo: 'M8 5 4 9l4 4 M4 9h8a4 4 0 0 1 0 8H9',
	redo: 'M12 5l4 4-4 4 M16 9H8a4 4 0 0 0 0 8h3',
	search: 'M9 4a5 5 0 1 0 0 10A5 5 0 0 0 9 4z M13 13l4 4',
	comment: 'M4 4h12v9H9l-4 3v-3H4z',
	back: 'M12 4 6 10l6 6',
	file: 'M6 3h6l4 4v10H6z M12 3v4h4',
	folder: 'M3 6h5l2 2h7v8H3z',
	copy: 'M7 7h9v10H7z M4 4h9v2 M4 4v10h2',
	print: 'M6 8V3h8v5 M4 8h12v6h-2 M6 12h8v5H6z',
	info: 'M10 3a7 7 0 1 0 0 14 7 7 0 0 0 0-14z M10 9v4 M10 6.5v.5',
	warning: 'M10 3 2 17h16z M10 8v4 M10 14.5v.5',
	pageView: 'M5 3h10v14H5z M8 7h4 M8 10h4',
	webView: 'M3 5h14v10H3z M3 8h14',
	minus: 'M5 10h10',
	plus: 'M5 10h10 M10 5v10',
} as const;

export type ChromeIcon = keyof typeof paths;

export function icon(name: ChromeIcon): SVGSVGElement {
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 20 20');
	svg.setAttribute('width', '16');
	svg.setAttribute('height', '16');
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('dve-icon');
	const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
	path.setAttribute('d', paths[name]);
	path.setAttribute('fill', 'none');
	path.setAttribute('stroke', 'currentColor');
	path.setAttribute('stroke-width', '1.5');
	path.setAttribute('stroke-linecap', 'round');
	path.setAttribute('stroke-linejoin', 'round');
	svg.append(path);
	return svg;
}

/** An icon button whose accessible name is `label` (localized later by `localizeElement`). */
export function iconButton(name: ChromeIcon, label: string, className = ''): HTMLButtonElement {
	const button = document.createElement('button');
	button.type = 'button';
	button.className = `dve-icon-button ${className}`.trim();
	button.setAttribute('aria-label', label);
	button.title = label;
	button.append(icon(name));
	return button;
}
