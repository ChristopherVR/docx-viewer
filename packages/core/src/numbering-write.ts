// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Paragraph } from './model.js';
import { children, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}

function sameNumbering(a: Paragraph['numbering'], b: Paragraph['numbering']): boolean {
	return a?.numId === b?.numId && a?.level === b?.level;
}

/** Writes or removes `w:numPr` on a paragraph's `pPr`, diffed against the base paragraph. */
export function writeNumberingProperties(
	doc: XmlDocument,
	props: XmlElement,
	paragraph: Paragraph,
	base?: Paragraph,
): void {
	if (base && sameNumbering(paragraph.numbering, base.numbering)) return;
	for (const existing of children(props, 'numPr')) props.removeChild(existing);
	if (!paragraph.numbering) return;
	const numPr = makeW(doc, 'numPr');
	const ilvl = makeW(doc, 'ilvl');
	setAttribute(ilvl, 'val', String(paragraph.numbering.level));
	numPr.appendChild(ilvl);
	const numId = makeW(doc, 'numId');
	setAttribute(numId, 'val', String(paragraph.numbering.numId));
	numPr.appendChild(numId);
	props.appendChild(numPr);
}
