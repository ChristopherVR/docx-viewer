// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, Note, Paragraph } from './model.js';
import { children, getW, parseXml, type XmlElement } from './xml.js';

const skippedTypes = new Set(['separator', 'continuationSeparator']);

/** Parses `word/footnotes.xml` or `word/endnotes.xml`, skipping Word's separator marks. */
export function parseNotesPart(
	xml: string,
	kind: 'footnote' | 'endnote',
	parseBlocks: (container: XmlElement, idPrefix: string) => Block[],
): Note[] {
	const document = parseXml(xml);
	const tag = kind === 'footnote' ? 'footnote' : 'endnote';
	const idPrefix = kind === 'footnote' ? 'fn' : 'en';
	const notes: Note[] = [];
	for (const element of children(document.documentElement, tag)) {
		const type = getW(element, 'type');
		if (type && skippedTypes.has(type)) continue;
		const id = getW(element, 'id');
		if (!id) continue;
		notes.push({ id, blocks: parseBlocks(element, `${idPrefix}${id}-`) });
	}
	return notes;
}

const romanNumerals: [number, string][] = [
	[1000, 'M'],
	[900, 'CM'],
	[500, 'D'],
	[400, 'CD'],
	[100, 'C'],
	[90, 'XC'],
	[50, 'L'],
	[40, 'XL'],
	[10, 'X'],
	[9, 'IX'],
	[5, 'V'],
	[4, 'IV'],
	[1, 'I'],
];
function toRoman(value: number): string {
	let remaining = Math.max(1, Math.round(value));
	let result = '';
	for (const [amount, numeral] of romanNumerals) {
		while (remaining >= amount) {
			result += numeral;
			remaining -= amount;
		}
	}
	return result || 'I';
}
function toLetters(value: number): string {
	let remaining = Math.max(1, Math.round(value));
	let result = '';
	while (remaining > 0) {
		remaining--;
		result = String.fromCharCode(97 + (remaining % 26)) + result;
		remaining = Math.floor(remaining / 26);
	}
	return result || 'a';
}

/** Renders a 1-based note number in a Word `w:numFmt` display format. */
export function formatNoteNumber(value: number, format: string | undefined): string {
	switch (format) {
		case 'upperRoman':
			return toRoman(value);
		case 'lowerRoman':
			return toRoman(value).toLowerCase();
		case 'upperLetter':
			return toLetters(value).toUpperCase();
		case 'lowerLetter':
			return toLetters(value);
		case 'decimalZero':
			return value < 10 ? `0${value}` : String(value);
		default:
			return String(value);
	}
}

function walkParagraph(
	paragraph: Paragraph,
	kind: 'footnote' | 'endnote',
	order: Map<string, number>,
) {
	for (const run of paragraph.runs) {
		if (run.noteReference?.kind === kind && !order.has(run.noteReference.id))
			order.set(run.noteReference.id, order.size + 1);
	}
}
function walkBlocks(blocks: Block[], kind: 'footnote' | 'endnote', order: Map<string, number>) {
	for (const block of blocks) {
		if (block.type === 'paragraph') walkParagraph(block, kind, order);
		else
			for (const row of block.rows)
				for (const cell of row) walkBlocks(cell.paragraphs, kind, order);
	}
}

/** Assigns 1-based note numbers by first-reference order, the same order Word displays. */
export function numberNotesInOrder(
	blocks: Block[],
	kind: 'footnote' | 'endnote',
): Map<string, number> {
	const order = new Map<string, number>();
	walkBlocks(blocks, kind, order);
	return order;
}
