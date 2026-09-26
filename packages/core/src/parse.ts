// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import JSZip from 'jszip';
import type {
	Block,
	DocumentModel,
	LoadedDocument,
	Paragraph,
	Table,
	TableCell,
	TextRun,
} from './model.js';
import {
	children,
	first,
	getW,
	isElement,
	named,
	parseXml,
	textContent,
	type XmlDocument,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { remember, saveDocx } from './save.js';
import { canEditTableStructure } from './write-table.js';
import { hasSpecialBreak } from './breaks.js';

const px = (twips: string | undefined, fallback: number): number =>
	twips === undefined ? fallback : (Number(twips) * 96) / 1440;
const points = (halfPoints: string | undefined): number | undefined =>
	halfPoints === undefined ? undefined : Number(halfPoints) / 2;
const twipValue = (value: string | undefined): number | undefined => {
	if (value === undefined || !/^-?\d+$/.test(value)) return undefined;
	const parsed = Number(value);
	return Number.isSafeInteger(parsed) ? parsed : undefined;
};
const on = (element: XmlElement | undefined): boolean => {
	if (!element) return false;
	const value = getW(element, 'val')?.toLowerCase();
	return !['0', 'false', 'off', 'no', 'none'].includes(value ?? '');
};

function parseRun(node: XmlElement): TextRun {
	const props = first(node, 'rPr');
	const text = Array.from(node.childNodes)
		.filter(isElement)
		.map((child) => {
			if (named(child, 't')) return textContent(child);
			if (named(child, 'tab')) return '\t';
			if (named(child, 'br') || named(child, 'cr')) return '\n';
			if (named(child, 'noBreakHyphen')) return '\u2011';
			return '';
		})
		.join('');
	const run: TextRun = { text };
	if (props && on(first(props, 'b'))) run.bold = true;
	if (props && on(first(props, 'i'))) run.italic = true;
	if (props && on(first(props, 'u'))) run.underline = true;
	if (props && on(first(props, 'strike') ?? first(props, 'dstrike'))) run.strike = true;
	const highlight = getW(first(props, 'highlight'), 'val');
	if (highlight) run.highlight = highlight;
	const verticalAlign = getW(first(props, 'vertAlign'), 'val');
	if (verticalAlign === 'superscript' || verticalAlign === 'subscript')
		run.verticalAlign = verticalAlign;
	const size = points(getW(first(props, 'sz'), 'val'));
	if (size !== undefined) run.fontSize = size;
	const fonts = first(props, 'rFonts');
	const family = getW(fonts, 'ascii') ?? getW(fonts, 'hAnsi');
	if (family) run.fontFamily = family;
	const hex = getW(first(props, 'color'), 'val');
	if (hex && /^[0-9a-f]{6}$/i.test(hex)) run.color = `#${hex}`;
	return run;
}

function parseParagraph(node: XmlElement, id: string): Paragraph {
	const props = first(node, 'pPr');
	const alignment = getW(first(props, 'jc'), 'val');
	const runs: TextRun[] = [];
	for (const item of Array.from(node.childNodes).filter(isElement)) {
		if (named(item, 'r')) runs.push(parseRun(item));
		else if (named(item, 'hyperlink'))
			Array.from(item.getElementsByTagNameNS(WORD_NS, 'r')).forEach((run: XmlElement) =>
				runs.push(parseRun(run)),
			);
	}
	if (!runs.length) runs.push({ text: '' });
	const paragraph: Paragraph = { type: 'paragraph', id, runs };
	if (
		alignment === 'left' ||
		alignment === 'center' ||
		alignment === 'right' ||
		alignment === 'both'
	)
		paragraph.align = alignment === 'both' ? 'justify' : alignment;
	const style = getW(first(props, 'pStyle'), 'val');
	if (style) paragraph.style = style;
	const spacing = first(props, 'spacing');
	const before = twipValue(getW(spacing, 'before'));
	const after = twipValue(getW(spacing, 'after'));
	const line = twipValue(getW(spacing, 'line'));
	if (before !== undefined) paragraph.spacingBeforeTwips = before;
	if (after !== undefined) paragraph.spacingAfterTwips = after;
	if (line !== undefined) {
		paragraph.lineSpacingTwips = line;
	}
	const rule = getW(spacing, 'lineRule');
	if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') paragraph.lineSpacingRule = rule;
	else if (line !== undefined) paragraph.lineSpacingRule = 'auto';
	const indent = first(props, 'ind');
	const indentLeft = twipValue(getW(indent, 'left'));
	const indentRight = twipValue(getW(indent, 'right'));
	const indentStart = twipValue(getW(indent, 'start'));
	const indentEnd = twipValue(getW(indent, 'end'));
	const firstLine = twipValue(getW(indent, 'firstLine'));
	const hanging = twipValue(getW(indent, 'hanging'));
	if (indentLeft !== undefined) paragraph.indentLeftTwips = indentLeft;
	if (indentRight !== undefined) paragraph.indentRightTwips = indentRight;
	if (indentStart !== undefined) paragraph.indentStartTwips = indentStart;
	if (indentEnd !== undefined) paragraph.indentEndTwips = indentEnd;
	if (firstLine !== undefined) paragraph.firstLineTwips = firstLine;
	if (hanging !== undefined) paragraph.hangingTwips = hanging;
	return paragraph;
}

function parseTable(node: XmlElement, id: string): Table {
	const rows = children(node, 'tr').map((row, ri) =>
		children(row, 'tc').map((cell, ci): TableCell => ({
			paragraphs: children(cell, 'p').map((p, pi) => parseParagraph(p, `${id}-r${ri}c${ci}p${pi}`)),
		})),
	);
	return { type: 'table', id, rows, structureEditable: canEditTableStructure(node) };
}

function hasAny(document: XmlDocument, names: string[]): boolean {
	return names.some((name) => document.getElementsByTagNameNS(WORD_NS, name).length > 0);
}

function warningsFor(document: XmlDocument): string[] {
	const warnings: string[] = [];
	const features: [string[], string][] = [
		[['drawing', 'pict', 'object'], 'Images and drawing objects are preserved but not editable.'],
		[['cols'], 'Multi-column layout is not represented in the document model.'],
		[
			['footnoteReference', 'endnoteReference'],
			'Footnotes and endnotes are not represented in the document model.',
		],
		[
			['commentRangeStart', 'trackRevisions'],
			'Comments and tracked review features are not represented in the document model.',
		],
		[['altChunk'], 'Embedded alternate-format content is not represented in the document model.'],
		[
			['hyperlink'],
			'Hyperlink targets are not represented in the document model; edits inside linked paragraphs are rejected to protect the original XML.',
		],
		[
			['numPr'],
			'List numbering is preserved as paragraph XML but is not represented in the document model.',
		],
		[
			['pStyle', 'rStyle'],
			'Style inheritance and theme font/color resolution are not modeled; displayed formatting may differ from Word.',
		],
	];
	for (const [names, message] of features) if (hasAny(document, names)) warnings.push(message);
	const specialBreak =
		Array.from(document.getElementsByTagNameNS(WORD_NS, 'br')).some(hasSpecialBreak) ||
		Array.from(document.getElementsByTagNameNS(WORD_NS, 'cr')).some(
			(cr: XmlElement) => cr.attributes.length > 0,
		);
	if (specialBreak)
		warnings.push(
			'Page, column, and other non-line breaks are not distinguished from line breaks in the document model; edits to paragraphs containing them are rejected to preserve the original XML.',
		);
	return warnings;
}

export interface PackageContext {
	original: Uint8Array;
	sourceXml: string;
	base: DocumentModel;
}

export async function readPackage(
	input: Uint8Array | ArrayBuffer,
): Promise<{ model: DocumentModel; context: PackageContext }> {
	const original =
		input instanceof Uint8Array ? new Uint8Array(input) : new Uint8Array(input.slice(0));
	if (original.byteLength > 50 * 1024 * 1024)
		throw new Error('DOCX package exceeds the 50 MiB compressed input limit');
	const zip = await JSZip.loadAsync(original);
	const entries = Object.values(zip.files);
	if (entries.length > 5000) throw new Error('DOCX package exceeds the 5,000 part limit');
	const totalSize = entries.reduce(
		(sum, entry) => sum + Number((entry as any)._data?.uncompressedSize ?? 0),
		0,
	);
	if (totalSize > 200 * 1024 * 1024)
		throw new Error('DOCX package exceeds the 200 MiB uncompressed content limit');
	const file = zip.file('word/document.xml');
	if (!file) throw new Error('DOCX package has no word/document.xml part');
	const sourceXml = await file.async('string');
	if (sourceXml.length > 25 * 1024 * 1024)
		throw new Error('word/document.xml exceeds the 25 MiB limit');
	const document: XmlDocument = parseXml(sourceXml);
	const body = Array.from(document.getElementsByTagNameNS(WORD_NS, 'body'))[0];
	if (!body) throw new Error('DOCX document.xml has no w:body');
	const blocks: Block[] = [];
	let index = 0;
	for (const node of Array.from(body.childNodes).filter(isElement)) {
		if (named(node, 'p')) blocks.push(parseParagraph(node, `p${index++}`));
		else if (named(node, 'tbl')) blocks.push(parseTable(node, `t${index++}`));
	}
	const section = children(body, 'sectPr').at(-1);
	const size = first(section, 'pgSz');
	const margins = first(section, 'pgMar');
	const model: DocumentModel = {
		blocks,
		page: {
			width: px(getW(size, 'w'), 816),
			height: px(getW(size, 'h'), 1056),
			marginTop: px(getW(margins, 'top'), 96),
			marginRight: px(getW(margins, 'right'), 96),
			marginBottom: px(getW(margins, 'bottom'), 96),
			marginLeft: px(getW(margins, 'left'), 96),
		},
		warnings: warningsFor(document),
	};
	if (zip.file('word/styles.xml'))
		model.warnings.push(
			'Word styles and theme inheritance are not resolved; inherited formatting can differ from Word.',
		);
	if (blocks.some((block) => block.type === 'table'))
		model.warnings.push(
			'Table text and cell structure are supported; table widths, borders, shading and cell formatting are not modeled.',
		);
	if (blocks.some((block) => block.type === 'table' && !block.structureEditable))
		model.warnings.push(
			'Merged, nested, or complex tables can be read, but their row and column structure cannot be edited safely.',
		);
	const context = { original, sourceXml, base: structuredClone(model) };
	remember(model, context);
	return { model, context };
}

export async function loadDocx(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	const { model, context } = await readPackage(input);
	return {
		model,
		save: (next = model) => {
			if (next !== model) remember(next, context);
			return saveDocx(next);
		},
	};
}
