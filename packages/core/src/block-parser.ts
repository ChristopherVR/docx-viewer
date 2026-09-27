// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Paragraph/table/run parsing shared by the main document body, headers, footers, footnotes and
// endnotes (moved out of parse.ts so every container can reuse the identical parser).
import type { Block, Paragraph, Table, TableCell, TextRun } from './model.js';
import {
	children,
	first,
	getW,
	isElement,
	named,
	textContent,
	type XmlElement,
	WORD_NS,
} from './xml.js';
import { canEditTableStructure } from './write-table.js';
import { classifyBreak } from './breaks.js';

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

const FIELD_PLACEHOLDERS: Record<string, string> = {
	PAGE: '[Page #]',
	NUMPAGES: '[Total pages]',
	DATE: '[Date]',
	TIME: '[Time]',
	FILENAME: '[File name]',
};
function fieldPlaceholderText(instr: string | undefined): string {
	const code = (instr ?? '').trim().split(/\s+/)[0]?.toUpperCase();
	return (code && FIELD_PLACEHOLDERS[code]) || '[Field]';
}

/** Content children of a run other than its `rPr`, used to detect single-purpose runs. */
function runContent(node: XmlElement): XmlElement[] {
	return Array.from(node.childNodes)
		.filter(isElement)
		.filter((child) => child.localName !== 'rPr');
}

function parseRun(node: XmlElement): TextRun {
	const props = first(node, 'rPr');
	const content = runContent(node);
	let breakKind: 'page' | 'column' | undefined;
	let noteReference: TextRun['noteReference'];
	if (content.length === 1) {
		const only = content[0];
		if (named(only, 'br')) {
			const kind = classifyBreak(only);
			if (kind === 'page' || kind === 'column') breakKind = kind;
		} else if (named(only, 'footnoteReference')) {
			const id = getW(only, 'id');
			if (id) noteReference = { kind: 'footnote', id };
		} else if (named(only, 'endnoteReference')) {
			const id = getW(only, 'id');
			if (id) noteReference = { kind: 'endnote', id };
		}
	}
	const text =
		breakKind || noteReference
			? ''
			: content
					.map((child) => {
						if (named(child, 't')) return textContent(child);
						if (named(child, 'tab')) return '\t';
						if (named(child, 'br') || named(child, 'cr')) return '\n';
						if (named(child, 'noBreakHyphen')) return '‑';
						return '';
					})
					.join('');
	const run: TextRun = { text };
	if (breakKind) run.break = breakKind;
	if (noteReference) run.noteReference = noteReference;
	const language = first(props, 'lang');
	const languageValue = getW(language, 'val');
	const eastAsiaLanguage = getW(language, 'eastAsia');
	const bidiLanguage = getW(language, 'bidi');
	if (languageValue !== undefined) run.language = languageValue;
	if (eastAsiaLanguage !== undefined) run.eastAsiaLanguage = eastAsiaLanguage;
	if (bidiLanguage !== undefined) run.bidiLanguage = bidiLanguage;
	const rtl = first(props, 'rtl');
	if (rtl) run.rtl = on(rtl);
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
		else if (named(item, 'fldSimple'))
			runs.push({ text: fieldPlaceholderText(getW(item, 'instr')) });
	}
	if (!runs.length) runs.push({ text: '' });
	const paragraph: Paragraph = { type: 'paragraph', id, runs };
	const bidi = first(props, 'bidi');
	if (bidi) paragraph.direction = on(bidi) ? 'rtl' : 'ltr';
	if (
		alignment === 'left' ||
		alignment === 'center' ||
		alignment === 'right' ||
		alignment === 'both'
	)
		paragraph.align = alignment === 'both' ? 'justify' : alignment;
	const style = getW(first(props, 'pStyle'), 'val');
	if (style) paragraph.style = style;
	if (on(first(props, 'pageBreakBefore'))) paragraph.pageBreakBefore = true;
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
	const numPr = first(props, 'numPr');
	const numId = twipValue(getW(first(numPr, 'numId'), 'val'));
	if (numPr && numId !== undefined && numId >= 0)
		paragraph.numbering = {
			numId,
			level: twipValue(getW(first(numPr, 'ilvl'), 'val')) ?? 0,
		};
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

/**
 * Parses every `w:p`/`w:tbl` direct child of a container (the document body, a header/footer
 * part's root, or a footnote/endnote element) into blocks, reusing one paragraph/table parser
 * everywhere. `idPrefix` namespaces ids so header/footer/note blocks never collide with the body.
 */
export function parseBlocksFromContainer(container: XmlElement, idPrefix = ''): Block[] {
	const blocks: Block[] = [];
	let index = 0;
	for (const node of Array.from(container.childNodes).filter(isElement)) {
		if (named(node, 'p')) blocks.push(parseParagraph(node, `${idPrefix}p${index++}`));
		else if (named(node, 'tbl')) blocks.push(parseTable(node, `${idPrefix}t${index++}`));
	}
	return blocks;
}

export { parseParagraph, parseTable };
