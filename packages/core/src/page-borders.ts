// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Page borders (`w:sectPr/w:pgBorders`): parsed into the section model and patched back side by
// side, so art borders and other attributes the model does not edit stay as they were.
import type { PageBorders, PageBorderSide } from './section-model.js';
import { borderSide } from './table-borders.js';
import { orderChildren } from './element-order.js';
import { parseUnsignedInteger } from './simple-types.js';
import { children, first, getW, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

const SIDES = ['top', 'left', 'bottom', 'right'] as const;
const ATTRIBUTES = ['offsetFrom', 'display', 'zOrder'] as const;
const setW = (element: XmlElement, name: string, value: string) =>
	element.setAttributeNS(WORD_NS, `w:${name}`, value);

export function parsePageBorders(element: XmlElement | undefined): PageBorders | undefined {
	if (!element) return undefined;
	const result: PageBorders = {};
	for (const side of SIDES) {
		const source = first(element, side);
		const value = borderSide(source);
		const art = getW(source, 'art');
		if (!source) continue;
		const space = parseUnsignedInteger(getW(source, 'space'));
		result[side] = {
			...(value ?? {}),
			...(space !== undefined ? { spacePoints: space } : {}),
			...(art ? { art } : {}),
		};
	}
	const offsetFrom = getW(element, 'offsetFrom');
	if (offsetFrom === 'page' || offsetFrom === 'text') result.offsetFrom = offsetFrom;
	const display = getW(element, 'display');
	if (display === 'allPages' || display === 'firstPage' || display === 'notFirstPage')
		result.display = display;
	const zOrder = getW(element, 'zOrder');
	if (zOrder === 'front' || zOrder === 'back') result.zOrder = zOrder;
	return result;
}

function buildSide(doc: XmlDocument, name: string, side: PageBorderSide): XmlElement {
	const edge = makeW(doc, name);
	setW(edge, 'val', side.style ?? 'single');
	setW(edge, 'sz', String(side.sizeEighthPoints ?? 4));
	setW(edge, 'space', String(side.spacePoints ?? 24));
	setW(edge, 'color', side.color ? side.color.replace(/^#/, '').toUpperCase() : 'auto');
	if (side.themeColor) setW(edge, 'themeColor', side.themeColor);
	if (side.art) setW(edge, 'art', side.art);
	return edge;
}

/**
 * Makes `sectPr`'s `w:pgBorders` match `next`, rewriting only the sides and attributes that differ
 * from `base`. `create` places a newly needed element in schema order.
 */
export function patchPageBorders(
	doc: XmlDocument,
	sectPr: XmlElement,
	next: PageBorders | undefined,
	base: PageBorders | undefined,
	create: (doc: XmlDocument, parent: XmlElement, name: string) => XmlElement,
): void {
	if (JSON.stringify(next ?? null) === JSON.stringify(base ?? null)) return;
	if (!next) {
		for (const old of children(sectPr, 'pgBorders')) sectPr.removeChild(old);
		return;
	}
	const element = create(doc, sectPr, 'pgBorders');
	for (const side of SIDES) {
		if (JSON.stringify(next[side]) === JSON.stringify(base?.[side])) continue;
		for (const old of children(element, side)) element.removeChild(old);
		const value = next[side];
		if (value) element.appendChild(buildSide(doc, side, value));
	}
	for (const name of ATTRIBUTES) {
		if (next[name] === base?.[name]) continue;
		if (next[name]) setW(element, name, next[name]);
		else element.removeAttributeNS(WORD_NS, name);
	}
	orderChildren(element, ['top', 'left', 'bottom', 'right']);
}
