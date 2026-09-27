// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Paragraph/table/run parsing shared by the main document body, headers, footers, footnotes and
// endnotes (moved out of parse.ts so every container can reuse the identical parser).
import type { Block, Paragraph, Revision, Table, TableCell, TextRun } from './model.js';
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
import { parseRunProperties } from './run-properties.js';
import { parseTable as parseTableWithFidelity } from './parse-table.js';
import { parseDrawing, type DrawingContext } from './drawing.js';
import { resolveHyperlink, parseSimpleHyperlinkField } from './hyperlink.js';
import { paragraphBookmarkNames } from './bookmarks.js';
import { createFieldTracker } from './field-runs.js';
import {
	collectParagraphRuns,
	paragraphFormatRevision,
	paragraphMarkRevision,
	runFormatRevision,
} from './parse-revisions.js';

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

/**
 * Package context (relationships, content types, media parts) for the container currently being
 * parsed. Only the main document body supplies one; headers, footers and notes have their own
 * relationship parts, so their drawings and external links are not resolved.
 */
let activeContext: DrawingContext | undefined;

/** Content children of a run other than its `rPr`, used to detect single-purpose runs. */
function runContent(node: XmlElement): XmlElement[] {
	return Array.from(node.childNodes)
		.filter(isElement)
		.filter((child) => child.localName !== 'rPr');
}

function parseRun(node: XmlElement, revision?: Revision): TextRun {
	const props = first(node, 'rPr');
	const content = runContent(node);
	const drawing = content.find((child) => named(child, 'drawing') || named(child, 'pict'));
	if (drawing && activeContext) {
		const run: TextRun = { text: '', image: parseDrawing(drawing, activeContext) };
		const runRevision = revision ?? runFormatRevision(props);
		if (runRevision) run.revision = runRevision;
		return run;
	}
	let breakKind: 'page' | 'column' | undefined;
	let noteReference: TextRun['noteReference'];
	let noteMark: TextRun['noteMark'];
	let fieldChar: TextRun['fieldChar'];
	const fieldCode =
		content.length > 0 && content.every((child) => named(child, 'instrText'))
			? content.map((child) => textContent(child)).join('')
			: undefined;
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
		} else if (named(only, 'footnoteRef')) noteMark = 'footnote';
		else if (named(only, 'endnoteRef')) noteMark = 'endnote';
		else if (named(only, 'fldChar')) {
			const type = getW(only, 'fldCharType');
			if (type === 'begin' || type === 'separate' || type === 'end') fieldChar = type;
		}
	}
	const text =
		breakKind || noteReference || noteMark || fieldChar || fieldCode !== undefined
			? ''
			: content
					.map((child) => {
						if (named(child, 't') || named(child, 'delText')) return textContent(child);
						if (named(child, 'tab')) return '\t';
						if (named(child, 'br') || named(child, 'cr')) return '\n';
						if (named(child, 'noBreakHyphen')) return '‑';
						return '';
					})
					.join('');
	const run: TextRun = { text, ...parseRunProperties(props) };
	if (breakKind) run.break = breakKind;
	if (noteReference) run.noteReference = noteReference;
	if (noteMark) run.noteMark = noteMark;
	if (fieldChar) run.fieldChar = fieldChar;
	if (fieldCode !== undefined) run.fieldCode = fieldCode;
	const runRevision = revision ?? runFormatRevision(props);
	if (runRevision) run.revision = runRevision;
	const language = first(props, 'lang');
	const languageValue = getW(language, 'val');
	const eastAsiaLanguage = getW(language, 'eastAsia');
	const bidiLanguage = getW(language, 'bidi');
	if (languageValue !== undefined) run.language = languageValue;
	if (eastAsiaLanguage !== undefined) run.eastAsiaLanguage = eastAsiaLanguage;
	if (bidiLanguage !== undefined) run.bidiLanguage = bidiLanguage;
	const rtl = first(props, 'rtl');
	if (rtl) run.rtl = on(rtl);
	const styleRef = getW(first(props, 'rStyle'), 'val');
	if (styleRef) run.style = styleRef;
	return run;
}

function parseParagraph(node: XmlElement, id: string): Paragraph {
	const props = first(node, 'pPr');
	const alignment = getW(first(props, 'jc'), 'val');
	const trackField = createFieldTracker();
	const { runs } = collectParagraphRuns(
		node,
		(element, revision) => trackField(element, parseRun(element, revision)),
		(item) => {
			if (!named(item, 'fldSimple')) return undefined;
			const instr = getW(item, 'instr') ?? '';
			const link = parseSimpleHyperlinkField(instr);
			const results = children(item, 'r').map((run) => parseRun(run));
			if (link) return results.map((run) => ({ ...run, link }));
			const field = { instr: instr.trim(), simple: true };
			// Show Word's cached result; fall back to a readable placeholder when none was saved.
			if (!results.some((run) => run.text)) return [{ text: fieldPlaceholderText(instr), field }];
			return results.map((run) => ({ ...run, field }));
		},
		(hyperlink) => {
			if (!activeContext) return undefined;
			const link = resolveHyperlink(hyperlink, activeContext.rels);
			return link.href !== undefined || link.anchor !== undefined ? link : undefined;
		},
	);
	if (!runs.length) runs.push({ text: '' });
	const paragraph: Paragraph = { type: 'paragraph', id, runs };
	const bookmarks = paragraphBookmarkNames(node);
	if (bookmarks.length) paragraph.bookmarks = bookmarks;
	const markRevision = paragraphMarkRevision(props);
	if (markRevision) paragraph.markRevision = markRevision;
	const formatRevision = paragraphFormatRevision(props);
	if (formatRevision) paragraph.formatRevision = formatRevision;
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
	return parseTableWithFidelity(node, id, parseParagraph);
}

/**
 * Parses every `w:p`/`w:tbl` direct child of a container (the document body, a header/footer
 * part's root, or a footnote/endnote element) into blocks, reusing one paragraph/table parser
 * everywhere. `idPrefix` namespaces ids so header/footer/note blocks never collide with the body.
 */
export function parseBlocksFromContainer(
	container: XmlElement,
	idPrefix = '',
	context?: DrawingContext,
): Block[] {
	const previous = activeContext;
	activeContext = context;
	try {
		return parseContainer(container, idPrefix);
	} finally {
		activeContext = previous;
	}
}

function parseContainer(container: XmlElement, idPrefix: string): Block[] {
	const blocks: Block[] = [];
	let index = 0;
	for (const node of Array.from(container.childNodes).filter(isElement)) {
		if (named(node, 'p')) blocks.push(parseParagraph(node, `${idPrefix}p${index++}`));
		else if (named(node, 'tbl')) blocks.push(parseTable(node, `${idPrefix}t${index++}`));
	}
	return blocks;
}

export { parseParagraph, parseTable };
