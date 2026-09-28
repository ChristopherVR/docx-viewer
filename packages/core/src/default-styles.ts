// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word's modern Normal-template defaults for documents created here. Without a styles.xml, Word
// falls back to its application defaults (Times New Roman 10pt, no spacing), so new documents
// would open looking different from how they were edited.
import type { DocumentModel } from './model.js';
import { parseRunStyleCatalog } from './character-styles.js';
import { parseParagraphStyleCatalog } from './paragraph-styles.js';
import { parseTableStyleCatalog } from './table-styles.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const heading = (
	id: string,
	name: string,
	level: number,
	before: number,
	size: number,
	color: string,
) =>
	`<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="0"/><w:outlineLvl w:val="${level}"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:eastAsia="Calibri Light" w:cs="Times New Roman"/><w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

const toc = (level: number) =>
	`<w:style w:type="paragraph" w:styleId="TOC${level}"><w:name w:val="toc ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="100"/>${level > 1 ? `<w:ind w:left="${(level - 1) * 220}"/>` : ''}</w:pPr></w:style>`;

/** Word's note text (10pt, single-spaced, no gap after) and superscript reference styles. */
const noteStyles = (id: string, name: string) =>
	`<w:style w:type="paragraph" w:styleId="${id}Text"><w:name w:val="${name} text"/><w:basedOn w:val="Normal"/><w:link w:val="${id}TextChar"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` +
	`<w:style w:type="character" w:styleId="${id}TextChar"><w:name w:val="${name} Text Char"/><w:basedOn w:val="DefaultParagraphFont"/><w:link w:val="${id}Text"/><w:uiPriority w:val="99"/><w:semiHidden/><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` +
	`<w:style w:type="character" w:styleId="${id}Reference"><w:name w:val="${name} reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`;

/** `word/styles.xml` for new documents: Word 2013+ docDefaults and its common built-in styles. */
export const DEFAULT_STYLES_XML =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W}">` +
	'<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
	'<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
	'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
	'<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
	'<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
	'<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
	'<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:eastAsia="Calibri Light" w:cs="Times New Roman"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>' +
	heading('Heading1', 'heading 1', 0, 240, 32, '2F5496') +
	heading('Heading2', 'heading 2', 1, 40, 26, '2F5496') +
	heading('Heading3', 'heading 3', 2, 40, 24, '1F3763') +
	toc(1) +
	toc(2) +
	toc(3) +
	noteStyles('Footnote', 'footnote') +
	noteStyles('Endnote', 'endnote') +
	'<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
	'</w:styles>';

let cached: Pick<DocumentModel, 'paragraphStyles' | 'characterStyles' | 'tableStyles'> | undefined;

/** Style catalogs parsed from `DEFAULT_STYLES_XML` (a fresh copy per call). */
export function defaultStyleCatalogs(): Required<
	Pick<DocumentModel, 'paragraphStyles' | 'characterStyles' | 'tableStyles'>
> {
	cached ??= {
		paragraphStyles: parseParagraphStyleCatalog(DEFAULT_STYLES_XML),
		characterStyles: parseRunStyleCatalog(DEFAULT_STYLES_XML),
		tableStyles: parseTableStyleCatalog(DEFAULT_STYLES_XML),
	};
	return structuredClone(cached) as Required<typeof cached>;
}

/** Whether `model` still carries exactly the built-in catalogs, so the default styles.xml fits it. */
export function hasDefaultStyles(model: DocumentModel): boolean {
	const defaults = defaultStyleCatalogs();
	return (
		JSON.stringify(model.paragraphStyles) === JSON.stringify(defaults.paragraphStyles) &&
		JSON.stringify(model.characterStyles ?? defaults.characterStyles) ===
			JSON.stringify(defaults.characterStyles) &&
		JSON.stringify(model.tableStyles ?? defaults.tableStyles) ===
			JSON.stringify(defaults.tableStyles)
	);
}
