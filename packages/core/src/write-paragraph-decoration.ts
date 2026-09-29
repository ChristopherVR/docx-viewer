// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes paragraph shading (`w:shd`) and borders (`w:pBdr`) when a paragraph's values changed.
// Unchanged paragraphs keep their source elements untouched.
import type { Paragraph } from './model.js';
import type { ParagraphBorders, TableBorderSide } from './table-model.js';
import { children, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

/** `CT_PBdr` child order (ECMA-376 §17.3.1.24). */
const SIDES = ['top', 'left', 'bottom', 'right', 'between'] as const;

/** A `w:color` value: six hex digits without `#`, or `auto`. */
function wordColor(color: string | undefined): string {
	const hex = /^#?([0-9a-f]{6})$/i.exec(color ?? '')?.[1];
	return hex ? hex.toUpperCase() : 'auto';
}

function borderElement(doc: XmlDocument, name: string, side: TableBorderSide): XmlElement {
	const element = makeW(doc, name);
	const set = (local: string, value: string) =>
		element.setAttributeNS(WORD_NS, `w:${local}`, value);
	set('val', side.style ?? 'single');
	if (side.style !== 'none' && side.style !== 'nil') {
		set('sz', String(side.sizeEighthPoints ?? 4));
		set('space', String(side.spacePoints ?? 1));
		set('color', wordColor(side.color));
		if (side.themeColor) set('themeColor', side.themeColor);
	}
	return element;
}

const sameJson = (a: unknown, b: unknown) =>
	JSON.stringify(a ?? null) === JSON.stringify(b ?? null);

function updateShading(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && paragraph.shadingFill === base.shadingFill) return;
	for (const element of children(props, 'shd')) props.removeChild(element);
	if (!paragraph.shadingFill) return;
	const shd = makeW(doc, 'shd');
	shd.setAttributeNS(WORD_NS, 'w:val', 'clear');
	shd.setAttributeNS(WORD_NS, 'w:color', 'auto');
	shd.setAttributeNS(WORD_NS, 'w:fill', wordColor(paragraph.shadingFill));
	props.appendChild(shd);
}

function updateBorders(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && sameJson(paragraph.borders, base.borders)) return;
	for (const element of children(props, 'pBdr')) props.removeChild(element);
	const borders: ParagraphBorders | undefined = paragraph.borders;
	if (!borders) return;
	const pBdr = makeW(doc, 'pBdr');
	for (const side of SIDES) {
		const value = borders[side];
		if (value) pBdr.appendChild(borderElement(doc, side, value));
	}
	if (pBdr.childNodes.length) props.appendChild(pBdr);
}

/** Applies changed shading and borders to `w:pPr`; the caller re-sorts the properties. */
export function writeParagraphDecoration(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	updateShading(doc, props, paragraph, base);
	updateBorders(doc, props, paragraph, base);
}
