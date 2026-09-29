// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes section edits: page size/orientation, margins, columns and break type for changed
// sections, new section breaks (a paragraph-level w:sectPr copied from the following section,
// as Word does) and removed breaks. Other section properties stay protected.
import type { Block, HeaderFooterSlots, SectionProperties } from './model.js';
import { orderSectionProperties } from './element-order.js';
import { orderParagraphProperties } from './tab-stops.js';
import { children, first, makeW, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

/** Section properties this writer can change; any other difference is rejected. */
const WRITABLE = new Set([
	'endsAtBlockId',
	'type',
	'pageWidthTwips',
	'pageHeightTwips',
	'orientation',
	'marginTopTwips',
	'marginRightTwips',
	'marginBottomTwips',
	'marginLeftTwips',
	'headerDistanceTwips',
	'footerDistanceTwips',
	'gutterTwips',
	'columns',
	'titlePage',
	'verticalAlign',
	'pageNumbering',
	'headers',
	'footers',
]);

function setW(element: XmlElement, name: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${name}`, value);
}

/** Returns the named child, creating it in the right place (schema order is honored by callers). */
function child(doc: XmlDocument, parent: XmlElement, name: string, before?: string[]): XmlElement {
	const existing = first(parent, name);
	if (existing) return existing;
	const element = makeW(doc, name);
	const anchor = before
		?.map((candidate) => first(parent, candidate))
		.find((candidate): candidate is XmlElement => Boolean(candidate));
	parent.insertBefore(element, anchor ?? null);
	return element;
}

// Elements that follow pgSz/pgMar/cols inside w:sectPr (CT_SectPr sequence).
const AFTER_PGSZ = [
	'pgMar',
	'paperSrc',
	'pgBorders',
	'lnNumType',
	'pgNumType',
	'cols',
	'formProt',
	'vAlign',
	'noEndnote',
	'titlePg',
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_COLS = [
	'formProt',
	'vAlign',
	'noEndnote',
	'titlePg',
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_PGNUMTYPE = ['cols', ...AFTER_COLS];
const AFTER_TITLEPG = [
	'textDirection',
	'bidi',
	'rtlGutter',
	'docGrid',
	'printerSettings',
	'sectPrChange',
];
const AFTER_TYPE = ['pgSz', ...AFTER_PGSZ];

function writeSectionProperties(doc: XmlDocument, sectPr: XmlElement, section: SectionProperties) {
	if (section.type === 'nextPage')
		for (const type of children(sectPr, 'type')) sectPr.removeChild(type);
	else setW(child(doc, sectPr, 'type', AFTER_TYPE), 'val', section.type);
	const size = child(doc, sectPr, 'pgSz', AFTER_PGSZ);
	setW(size, 'w', String(section.pageWidthTwips));
	setW(size, 'h', String(section.pageHeightTwips));
	if (section.orientation === 'landscape') setW(size, 'orient', 'landscape');
	else size.removeAttributeNS(WORD_NS, 'orient');
	const margins = child(doc, sectPr, 'pgMar', AFTER_PGSZ.slice(1));
	setW(margins, 'top', String(section.marginTopTwips));
	setW(margins, 'right', String(section.marginRightTwips));
	setW(margins, 'bottom', String(section.marginBottomTwips));
	setW(margins, 'left', String(section.marginLeftTwips));
	setW(margins, 'header', String(section.headerDistanceTwips ?? 720));
	setW(margins, 'footer', String(section.footerDistanceTwips ?? 720));
	setW(margins, 'gutter', String(section.gutterTwips ?? 0));
	const columns = child(doc, sectPr, 'cols', AFTER_COLS);
	setW(columns, 'num', String(section.columns.count));
	setW(columns, 'space', String(section.columns.spacingTwips ?? 720));
	if (section.columns.equalWidth || !section.columns.widths?.length) {
		for (const column of children(columns, 'col')) columns.removeChild(column);
		columns.removeAttributeNS(WORD_NS, 'equalWidth');
	}
	if (section.columns.separator) setW(columns, 'sep', '1');
	else columns.removeAttributeNS(WORD_NS, 'sep');
	if (section.verticalAlign && section.verticalAlign !== 'top')
		setW(
			child(doc, sectPr, 'vAlign', ['noEndnote', 'titlePg', ...AFTER_TITLEPG]),
			'val',
			section.verticalAlign,
		);
	else for (const vAlign of children(sectPr, 'vAlign')) sectPr.removeChild(vAlign);
	if (section.titlePage) child(doc, sectPr, 'titlePg', AFTER_TITLEPG);
	else for (const titlePg of children(sectPr, 'titlePg')) sectPr.removeChild(titlePg);
	const numbering = section.pageNumbering;
	if (numbering?.format || numbering?.start !== undefined) {
		const pgNumType = child(doc, sectPr, 'pgNumType', AFTER_PGNUMTYPE);
		if (numbering.format && numbering.format !== 'decimal')
			setW(pgNumType, 'fmt', numbering.format);
		else pgNumType.removeAttributeNS(WORD_NS, 'fmt');
		if (numbering.start !== undefined) setW(pgNumType, 'start', String(numbering.start));
		else pgNumType.removeAttributeNS(WORD_NS, 'start');
	} else
		for (const pgNumType of children(sectPr, 'pgNumType')) {
			pgNumType.removeAttributeNS(WORD_NS, 'fmt');
			pgNumType.removeAttributeNS(WORD_NS, 'start');
			if (!pgNumType.attributes.length) sectPr.removeChild(pgNumType);
		}
}

function assertWritable(section: SectionProperties, base: SectionProperties | undefined) {
	if (!base) return;
	for (const key of new Set([...Object.keys(section), ...Object.keys(base)]))
		if (
			!WRITABLE.has(key) &&
			JSON.stringify(section[key as keyof SectionProperties]) !==
				JSON.stringify(base[key as keyof SectionProperties])
		)
			throw new Error(
				`Editing the section property "${key}" is not supported; only page size, orientation, margins, columns, break type, vertical alignment, first-page setting and page numbering can change.`,
			);
}

const REL_NAMESPACE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * Adds `w:headerReference` / `w:footerReference` for header and footer parts this section gained
 * (`references` maps each new part to its relationship id). Existing references stay as they are.
 */
function writeNewReferences(
	doc: XmlDocument,
	sectPr: XmlElement,
	section: SectionProperties,
	previous: SectionProperties | undefined,
	references: ReadonlyMap<string, string>,
): void {
	for (const [kind, name] of [
		['headers', 'headerReference'],
		['footers', 'footerReference'],
	] as const)
		for (const [slot, content] of Object.entries(section[kind] ?? {}) as Array<
			[keyof HeaderFooterSlots, HeaderFooterSlots[keyof HeaderFooterSlots]]
		>) {
			const id = content?.partName ? references.get(content.partName) : undefined;
			if (!id || previous?.[kind]?.[slot]?.partName === content?.partName) continue;
			for (const old of children(sectPr, name))
				if (old.getAttributeNS(WORD_NS, 'type') === slot) sectPr.removeChild(old);
			const reference = makeW(doc, name);
			setW(reference, 'type', slot);
			reference.setAttributeNS(REL_NAMESPACE, 'r:id', id);
			sectPr.appendChild(reference);
		}
}

/** The paragraph-level w:sectPr of a top-level paragraph, if any. */
const paragraphSection = (node: XmlElement | undefined) =>
	node?.localName === 'p' ? first(first(node, 'pPr'), 'sectPr') : undefined;

/**
 * Applies `sections` to a body already rewritten for `blocks`. `base` is the loaded section list
 * (empty for new documents); unchanged sections are left byte-identical.
 */
export function applySectionEdits(
	doc: XmlDocument,
	body: XmlElement,
	blocks: Block[],
	sections: SectionProperties[],
	base: SectionProperties[],
	/** Relationship ids of header/footer parts created for this save, by part name. */
	references: ReadonlyMap<string, string> = new Map(),
): void {
	if (!sections.length) {
		if (base.length)
			throw new Error('A document needs at least one section; the section list cannot be emptied.');
		return;
	}
	const topLevel = Array.from(body.childNodes).filter(
		(node): node is XmlElement =>
			node.nodeType === 1 && ['p', 'tbl'].includes((node as XmlElement).localName),
	);
	const nodeFor = (id: string) => topLevel[blocks.findIndex((block) => block.id === id)];
	const ends = new Set(sections.slice(0, -1).map((section) => section.endsAtBlockId));
	// Remove breaks the model no longer has.
	blocks.forEach((block, index) => {
		const sectPr = paragraphSection(topLevel[index]);
		if (sectPr && !ends.has(block.id)) sectPr.parentNode?.removeChild(sectPr);
	});
	let bodySection = children(body, 'sectPr').at(-1);
	if (!bodySection) {
		bodySection = makeW(doc, 'sectPr');
		body.appendChild(bodySection);
	}
	const baseByEnd = new Map(base.map((section) => [section.endsAtBlockId, section]));
	// Walk backwards so each new break can copy the section that follows it.
	let following = bodySection;
	for (let index = sections.length - 1; index >= 0; index--) {
		const section = sections[index];
		if (!section) continue;
		const last = index === sections.length - 1;
		let sectPr: XmlElement | undefined = last
			? bodySection
			: paragraphSection(nodeFor(section.endsAtBlockId));
		const previous = last ? base.at(-1) : baseByEnd.get(section.endsAtBlockId);
		if (!sectPr) {
			const paragraph = nodeFor(section.endsAtBlockId);
			if (paragraph?.localName !== 'p')
				throw new Error('A section break must end on a paragraph, not a table.');
			let pPr = first(paragraph, 'pPr');
			if (!pPr) {
				pPr = makeW(doc, 'pPr');
				paragraph.insertBefore(pPr, paragraph.firstChild);
			}
			sectPr = following.cloneNode(true) as XmlElement;
			pPr.appendChild(sectPr);
			orderParagraphProperties(pPr);
		}
		assertWritable(section, previous);
		if (!previous || JSON.stringify(previous) !== JSON.stringify(section)) {
			writeSectionProperties(doc, sectPr, section);
			writeNewReferences(doc, sectPr, section, previous, references);
			orderSectionProperties(sectPr);
		}
		following = sectPr;
	}
}
