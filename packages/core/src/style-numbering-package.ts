// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type JSZip from 'jszip';
import type { DocumentModel } from './model.js';
import type { ParagraphStyleCatalog } from './model-paragraph.js';
import { orderParagraphProperties } from './tab-stops.js';
import {
	buildXml,
	children,
	first,
	getW,
	makeW,
	parseXml,
	WORD_NS,
	type XmlElement,
} from './xml.js';

const STYLES_PART = 'word/styles.xml';
/** Children of `w:style` that follow `w:pPr` in the schema sequence. */
const AFTER_PPR = ['rPr', 'tblPr', 'trPr', 'tcPr', 'tblStylePr'];

const withoutNumbering = (catalog: ParagraphStyleCatalog | undefined) =>
	catalog
		? {
				...catalog,
				styles: Object.fromEntries(
					Object.entries(catalog.styles).map(([id, style]) => {
						const { numbering: _numbering, ...rest } = style;
						return [id, rest];
					}),
				),
			}
		: undefined;

/** Whether the only difference between the two style catalogs is style-level `numbering`. */
export function differsOnlyByStyleNumbering(
	next: ParagraphStyleCatalog | undefined,
	prior: ParagraphStyleCatalog | undefined,
): boolean {
	return JSON.stringify(withoutNumbering(next)) === JSON.stringify(withoutNumbering(prior));
}

function setNumPr(style: XmlElement, numbering: { numId: number; level: number } | undefined) {
	const doc = style.ownerDocument;
	let pPr = first(style, 'pPr');
	if (!numbering && !pPr) return;
	if (!pPr) {
		pPr = makeW(doc, 'pPr');
		const anchor = elementsOf(style).find((child) => AFTER_PPR.includes(child.localName)) ?? null;
		style.insertBefore(pPr, anchor);
	}
	for (const existing of children(pPr, 'numPr')) pPr.removeChild(existing);
	if (numbering) {
		const numPr = makeW(doc, 'numPr');
		const level = makeW(doc, 'ilvl');
		level.setAttributeNS(WORD_NS, 'w:val', String(numbering.level));
		const id = makeW(doc, 'numId');
		id.setAttributeNS(WORD_NS, 'w:val', String(numbering.numId));
		numPr.appendChild(level);
		numPr.appendChild(id);
		pPr.appendChild(numPr);
	}
	orderParagraphProperties(pPr);
}
const elementsOf = (parent: XmlElement) =>
	Array.from(parent.childNodes).filter((node): node is XmlElement => node.nodeType === 1);

/**
 * Writes style-level `numPr` changes (heading-linked lists) into the existing `word/styles.xml`.
 * Only `numbering` of styles already in the package may change; everything else is preserved.
 */
export async function applyStyleNumbering(
	zip: JSZip,
	model: DocumentModel,
	base: DocumentModel,
): Promise<void> {
	const next = model.paragraphStyles?.styles ?? {};
	const prior = base.paragraphStyles?.styles ?? {};
	const changed = Object.keys(next).filter(
		(id) => JSON.stringify(next[id]?.numbering) !== JSON.stringify(prior[id]?.numbering),
	);
	if (!changed.length) return;
	const file = zip.file(STYLES_PART);
	if (!file) throw new Error('Cannot link styles to a list: the package has no styles.xml.');
	const doc = parseXml(await file.async('string'));
	const styles = children(doc.documentElement, 'style');
	for (const id of changed) {
		const style = styles.find(
			(element) => getW(element, 'type') === 'paragraph' && getW(element, 'styleId') === id,
		);
		if (!style) throw new Error(`Cannot link style "${id}" to a list: it is not in styles.xml.`);
		setNumPr(style, next[id]?.numbering);
	}
	zip.file(STYLES_PART, buildXml(doc));
}
