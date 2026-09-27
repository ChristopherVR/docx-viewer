// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes section edits: page size/orientation, margins, columns and break type for changed
// sections, new section breaks (a paragraph-level w:sectPr copied from the following section,
// as Word does) and removed breaks. Other section properties stay protected.
import type { Block, SectionProperties } from './model.js';
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
				`Editing the section property "${key}" is not supported; only page size, orientation, margins, columns and break type can change.`,
			);
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
		}
		assertWritable(section, previous);
		if (!previous || JSON.stringify(previous) !== JSON.stringify(section))
			writeSectionProperties(doc, sectPr, section);
		following = sectPr;
	}
}
