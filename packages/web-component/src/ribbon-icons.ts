/** Line icons for ribbon buttons on a 24x24 grid, drawn with `currentColor` so themes apply. */
const paths = {
	paste: 'M8 4h8v3H8z M6 5.5H5v15h14v-15h-1 M9 12h6 M9 15.5h6',
	cut: 'M6 3l9 12 M18 3l-9 12 M7 15.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z M17 15.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z',
	copy: 'M9 8h11v13H9z M6 16H4V3h11v2',
	bold: 'M7 4h6a3.5 3.5 0 0 1 0 7H7z M7 11h7a3.5 3.5 0 0 1 0 7H7z',
	italic: 'M10 4h8 M6 20h8 M15 4l-6 16',
	underline: 'M7 4v7a5 5 0 0 0 10 0V4 M5 21h14',
	strike:
		'M4 12h16 M16.5 7.5C16 5.5 14 4.5 12 4.5c-2.5 0-4.2 1.3-4.2 3.2 0 1.6 1.2 2.4 4.2 3 M8 16.5c.5 2 2.3 3 4.2 3 2.6 0 4.3-1.3 4.3-3.3 0-1.2-.6-2-2-2.5',
	subscript: 'M4 5l8 10 M12 5L4 15 M16 17.5a2 2 0 1 1 3.5 1.3L16 22h4',
	superscript: 'M4 9l8 10 M12 9L4 19 M16 3.5a2 2 0 1 1 3.5 1.3L16 8h4',
	fontGrow: 'M3 19 8.5 5 14 19 M5 14h7 M17 9V3 M14 6h6',
	fontShrink: 'M3 19 8.5 5 14 19 M5 14h7 M14 6h6',
	clear: 'M4 20l4-1 11-11-3-3L5 16z M13 7l3 3 M14 20h7',
	fontColor: 'M6 17 12 4l6 13 M8.3 12h7.4',
	highlight: 'M14 4l6 6-8 8H8l-4-4v-3z M4 21h9',
	bullets: 'M9 6h11 M9 12h11 M9 18h11 M4.5 6h.01 M4.5 12h.01 M4.5 18h.01',
	numbering:
		'M11 6h9 M11 12h9 M11 18h9 M4 4.5l1.5-1V9 M4 14.5c.5-1.5 3-1.5 3 .3 0 1-1.5 1.6-3 3.2h3',
	outdent: 'M8 6h12 M12 11h8 M12 16h8 M8 21h12 M7 9l-4 3 4 3z',
	indent: 'M8 6h12 M12 11h8 M12 16h8 M8 21h12 M3 9l4 3-4 3z',
	removeList: 'M9 6h11 M9 12h6 M9 18h11 M4 4l5 5 M9 4 4 9',
	alignLeft: 'M4 6h16 M4 10h10 M4 14h16 M4 18h10',
	alignCenter: 'M4 6h16 M7 10h10 M4 14h16 M7 18h10',
	alignRight: 'M4 6h16 M10 10h10 M4 14h16 M10 18h10',
	justify: 'M4 6h16 M4 10h16 M4 14h16 M4 18h16',
	lineSpacing: 'M10 6h10 M10 12h10 M10 18h10 M5 4v16 M3 6l2-2 2 2 M3 18l2 2 2-2',
	spaceBefore: 'M4 20h16 M8 4v10 M5.5 11.5 8 14l2.5-2.5',
	spaceAfter: 'M4 4h16 M8 20V10 M5.5 12.5 8 10l2.5 2.5',
	find: 'M10 4a6 6 0 1 0 0 12 6 6 0 0 0 0-12z M14.5 14.5 20 20',
	table: 'M4 5h16v14H4z M4 10h16 M4 15h16 M10 5v14 M15 5v14',
	picture: 'M4 5h16v14H4z M4 16l5-5 4 4 3-3 4 4 M15.5 9.5h.01',
	formatPicture: 'M4 6h12v12H4z M4 15l3-3 3 3 M16 4l4 4-7 7h-4v-4z',
	link: 'M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1 M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1',
	pageBreak: 'M7 3h10v6H7z M7 15h10v6H7z M3 12h3 M9 12h2 M13 12h2 M18 12h3',
	columnBreak: 'M4 4h6v16H4z M14 4h6v16h-6z M12 4v3 M12 10v4 M12 17v3',
	toc: 'M4 6h16 M8 11h12 M8 16h12 M4 11h.01 M4 16h.01 M4 21h16',
	updateTable: 'M20 12a8 8 0 1 1-2.5-5.8 M20 4v5h-5',
	footnote: 'M5 4h9v12H5z M8 8h3 M8 11h3 M4 20h16 M17 6v.01',
	endnote: 'M5 3h9v11H5z M8 7h3 M8 10h3 M4 17h16 M4 21h16',
	sectionNext: 'M5 3h9l3 3v8H5z M5 18h14 M5 21h14 M12 9v4 M10 11l2 2 2-2',
	sectionContinuous: 'M5 4h14v6H5z M5 14h14v6H5z M3 12h2 M8 12h2 M14 12h2 M19 12h2',
	sectionEven: 'M6 3h9l3 3v15H6z M9 13a1.5 1.5 0 0 1 3 0c0 1.2-3 2.5-3 3.5h3',
	sectionOdd: 'M6 3h9l3 3v15H6z M9 12l1.5-1v6',
	firstPage: 'M6 3h9l3 3v15H6z M9 8h6 M9 12h6 M9 16h3',
	oddEven: 'M3 5h8v14H3z M13 5h8v14h-8z M6 9h2 M16 9h2',
	track: 'M4 20l4-1 11-11-3-3L5 16z M14 20h7',
	previous: 'M6 14l6-6 6 6',
	next: 'M6 10l6 6 6-6',
	accept: 'M5 12.5l4.5 4.5L19 7',
	reject: 'M6 6l12 12 M18 6L6 18',
	acceptAll: 'M2 12.5l4 4L14 8 M10 15.5l1.5 1.5L21 7',
	rejectAll: 'M4 6l7 7 M11 6l-7 7 M14 12l7 7 M21 12l-7 7',
	comment: 'M4 5h16v11H10l-4 4v-4H4z M8 9h8 M8 12h5',
	comments: 'M3 4h13v9H9l-3 3v-3H3z M9 17h6l3 3v-3h3V8h-3',
	hidden:
		'M3 12s3.5-6 9-6 9 6 9 6-3.5 6-9 6-9-6-9-6z M12 9.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z M4 20L20 4',
	thumbnails: 'M4 4h6v7H4z M14 4h6v7h-6z M4 14h6v7H4z M14 14h6v7h-6z',
	print: 'M7 9V3h10v6 M5 9h14v8h-3 M7 14h10v7H7z',
	rowAbove: 'M4 10h16v10H4z M4 15h16 M12 3v4 M10 5h4',
	rowBelow: 'M4 4h16v10H4z M4 9h16 M12 17v4 M10 19h4',
	deleteRow: 'M4 6h16v9H4z M4 10.5h16 M8 19l8 0',
	columnLeft: 'M10 4h10v16H10z M15 4v16 M3 12h4 M5 10v4',
	columnRight: 'M4 4h10v16H4z M9 4v16 M17 12h4 M19 10v4',
	deleteColumn: 'M6 4h9v16H6z M10.5 4v16 M19 8v8',
	deleteTable: 'M4 5h16v14H4z M4 10h16 M10 5v14 M14 13l5 5 M19 13l-5 5',
	margins: 'M5 3h14v18H5z M8 6h8v12H8z',
	orientation: 'M6 3h9l3 3v15H6z M12 9v6 M9 12h6',
	columns: 'M4 4h7v16H4z M13 4h7v16h-7z M6 8h3 M6 12h3 M15 8h3 M15 12h3',
	verticalAlign: 'M5 4h14 M5 20h14 M8 8h8 M8 12h8 M8 16h8',
	pageNumber: 'M6 3h9l3 3v15H6z M9 17h6 M12 11v4 M10.5 12.5l1.5-1.5',
	numbering2: 'M6 3h9l3 3v15H6z M9 9h6 M9 13h6 M9 17h3',
	view: 'M3 12s3.5-6 9-6 9 6 9 6-3.5 6-9 6-9-6-9-6z M12 9.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5z',
	zoom: 'M10 4a6 6 0 1 0 0 12 6 6 0 0 0 0-12z M14.5 14.5 20 20 M7.5 10h5 M10 7.5v5',
	markup: 'M4 5h16v11H10l-4 4v-4H4z M8 9h8 M8 12h4',
	changeCase: 'M2 18 7 5l5 13 M4 14h6 M14 18l3.5-8 3.5 8 M15.5 15.5h4',
	formatPainter: 'M5 4h11v5H5z M16 6.5h3v5h-8v3 M9.5 14.5h3V21h-3z',
	paragraphMarks: 'M12 4H9a4 4 0 0 0 0 8h3 M12 4v16 M16 4v16',
	select: 'M6 3l12 9-5.5 1.2L15 19l-2.5 1-2.5-5.8L6 18z',
	replace: 'M4 7h12l-3-3 M20 17H8l3 3 M10 12a3 3 0 1 0 0 .01',
	moreStyles: 'M4 5h16 M4 10h16 M4 15h10 M17 16l3 3 3-3',
	charStyle: 'M4 19 9 5l5 14 M6 14h6 M17 9h4 M19 9v10',
	wordCount: 'M4 5h16v14H4z M8 15V9l2 4 2-4v6 M16 9v6',
	spelling: 'M4 15l4-10 4 10 M5.5 12h5 M14 17l2.5 2.5L21 13',
	symbol: 'M5 19h4v-2a6 6 0 1 1 6 0v2h4',
	dateTime: 'M4 5h16v15H4z M4 10h16 M8 3v4 M16 3v4 M12 13v3l2 1',
	zoom100: 'M10 4a6 6 0 1 0 0 12 6 6 0 0 0 0-12z M14.5 14.5 20 20 M8.5 8.5l1.5-1v5',
	pageWidth: 'M3 6v12 M21 6v12 M7 12h10 M9 10l-2 2 2 2 M15 10l2 2-2 2',
	onePage: 'M7 3h10v18H7z M10 8h4 M10 12h4',
	launcher: 'M14 4h6v6 M20 4l-7 7 M10 6H5v13h13v-5',
	pageSize: 'M6 3h9l3 3v15H6z M9 11h6 M9 15h6 M4 9v8 M2.5 10.5 4 9l1.5 1.5',
	header: 'M5 3h14v18H5z M5 8h14 M8 5.5h6',
	footer: 'M5 3h14v18H5z M5 16h14 M8 18.5h6',
	pageNumberIcon: 'M5 3h14v18H5z M9 17h6 M12 11v4 M10.5 12.5l1.5-1',
	shading:
		'M5 11l6-6 8 8-6 6z M11 5l-3-2 M19 16c1 1.5 1.5 2.5 1.5 3.3a1.5 1.5 0 0 1-3 0c0-.8.5-1.8 1.5-3.3z',
	borders: 'M4 4h16v16H4z M4 12h16 M12 4v16',
	addText: 'M4 5h16 M4 10h10 M4 15h16 M17 12v6 M14 15h6',
	blankPage: 'M6 3h9l3 3v15H6z',
	readAloud: 'M4 9v6h4l5 4V5L8 9z M16 9a4 4 0 0 1 0 6 M18.5 6.5a8 8 0 0 1 0 11',
	gridlines: 'M4 4h16v16H4z M4 9h16 M4 14h16 M9 4v16 M14 4v16',
	caret: 'M7 10l5 5 5-5',
} as const;

export type RibbonIcon = keyof typeof paths;

export function isRibbonIcon(name: string): name is RibbonIcon {
	return name in paths;
}

/** An `<svg>` for `name`; `size` is its rendered edge in CSS pixels. */
export function ribbonIcon(name: RibbonIcon, size = 16): SVGSVGElement {
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	svg.setAttribute('viewBox', '0 0 24 24');
	svg.setAttribute('width', String(size));
	svg.setAttribute('height', String(size));
	svg.setAttribute('aria-hidden', 'true');
	svg.classList.add('ribbon-icon');
	const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
	path.setAttribute('d', paths[name]);
	path.setAttribute('fill', 'none');
	path.setAttribute('stroke', 'currentColor');
	path.setAttribute('stroke-width', '1.7');
	path.setAttribute('stroke-linecap', 'round');
	path.setAttribute('stroke-linejoin', 'round');
	svg.append(path);
	return svg;
}
