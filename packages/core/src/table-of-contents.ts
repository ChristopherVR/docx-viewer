// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Table of contents fields (`TOC \o "1-3"`): builds entries from heading paragraphs and finds or
// refreshes an existing TOC field that spans body paragraphs.
import { fieldName } from './field-runs.js';
import type { Block, DocumentModel, Paragraph, ParagraphStyleCatalog, TextRun } from './model.js';

export const DEFAULT_TOC_INSTRUCTION = ' TOC \\o "1-3" \\h \\z \\u ';
/** Word's result text when a TOC has no entries. */
export const EMPTY_TOC_TEXT = 'No table of contents entries found.';

/** The heading level (1–9) of a paragraph, from its style id or Word's built-in `heading N` name. */
export function headingLevel(
	paragraph: Paragraph,
	catalog?: ParagraphStyleCatalog,
): number | undefined {
	const seen = new Set<string>();
	let styleId = paragraph.style;
	while (styleId && !seen.has(styleId)) {
		seen.add(styleId);
		const style = catalog?.styles[styleId];
		const match = /^heading\s*([1-9])$/i.exec(style?.name ?? styleId);
		if (match) return Number(match[1]);
		styleId = style?.basedOn;
	}
	return undefined;
}

/** The outline levels a TOC instruction collects (`\o "1-3"`); Word's default is 1–9. */
export function tocLevels(instruction: string): { from: number; to: number } {
	const match = /\\o\s*"?(\d)\s*-\s*(\d)"?/.exec(instruction);
	return match ? { from: Number(match[1]), to: Number(match[2]) } : { from: 1, to: 9 };
}

const plainText = (paragraph: Paragraph) =>
	paragraph.runs
		.filter((run) => !run.fieldCode && !run.fieldChar && !run.noteReference && !run.image)
		.filter((run) => run.revision?.kind !== 'delete' && run.revision?.kind !== 'moveFrom')
		.map((run) => run.text.replace(/[\t\n]/g, ' '))
		.join('')
		.trim();

export interface TocEntry {
	level: number;
	text: string;
	blockId: string;
}

/** Heading paragraphs in body order that a TOC with `instruction` lists. */
export function tocEntries(
	model: DocumentModel,
	instruction = DEFAULT_TOC_INSTRUCTION,
): TocEntry[] {
	const { from, to } = tocLevels(instruction);
	const entries: TocEntry[] = [];
	for (const block of model.blocks) {
		if (block.type !== 'paragraph') continue;
		const level = headingLevel(block, model.paragraphStyles);
		const text = plainText(block);
		if (level !== undefined && level >= from && level <= to && text)
			entries.push({ level, text, blockId: block.id });
	}
	return entries;
}

export interface TocOptions {
	instruction?: string;
	/** Page number text per heading block id, from pagination; entries without one show no number. */
	pageNumbers?: ReadonlyMap<string, string>;
	/** Text width in twips, for the right-aligned page-number tab stop. */
	contentWidthTwips?: number;
	/** Creates unique paragraph ids. */
	newId: () => string;
	/**
	 * `_Toc` bookmark per heading block id. With it, entries link to their heading and number
	 * their page with a `PAGEREF` field, as Word writes a TOC with the `\h` switch.
	 */
	bookmarks?: ReadonlyMap<string, string>;
}

/** An entry as Word writes it: a hyperlink to the heading's bookmark holding text, tab and PAGEREF. */
function linkedEntry(text: string, bookmark: string, page: string | undefined): TextRun[] {
	const link = { anchor: bookmark };
	const runs: TextRun[] = [{ text: page === undefined ? text : `${text}\t`, link }];
	if (page === undefined) return runs;
	const instr = `PAGEREF ${bookmark} \\h`;
	return [
		...runs,
		{ text: '', fieldChar: 'begin', link },
		{ text: '', fieldCode: ` ${instr} `, link },
		{ text: '', fieldChar: 'separate', link },
		{ text: page, field: { instr }, link },
		{ text: '', fieldChar: 'end', link },
	];
}

/**
 * The `_Toc` bookmark each entry links to: the heading's existing one, or a new unique name.
 * Returns only headings that need a new bookmark added in `added`.
 */
export function tocBookmarks(
	model: DocumentModel,
	entries: TocEntry[],
): { bookmarks: Map<string, string>; added: Map<string, string> } {
	const used = new Set<string>();
	const paragraphs = new Map<string, Paragraph>();
	for (const block of model.blocks)
		for (const paragraph of block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs))) {
			paragraphs.set(paragraph.id, paragraph);
			for (const name of paragraph.bookmarks ?? []) used.add(name);
		}
	const bookmarks = new Map<string, string>();
	const added = new Map<string, string>();
	let seed = 100000000;
	for (const entry of entries) {
		const existing = paragraphs
			.get(entry.blockId)
			?.bookmarks?.find((name) => name.startsWith('_Toc'));
		if (existing) {
			bookmarks.set(entry.blockId, existing);
			continue;
		}
		while (used.has(`_Toc${seed}`)) seed++;
		const name = `_Toc${seed++}`;
		used.add(name);
		bookmarks.set(entry.blockId, name);
		added.set(entry.blockId, name);
	}
	return { bookmarks, added };
}

/** Whether a TOC instruction asks for hyperlinked entries (`\h`). */
export const tocHyperlinks = (instruction: string): boolean => /\\h\b/.test(instruction);

/** `model` with new `_Toc` bookmarks added to the heading paragraphs named in `added`. */
export function withTocBookmarks(
	model: DocumentModel,
	added: ReadonlyMap<string, string>,
): DocumentModel {
	if (!added.size) return model;
	const mark = (paragraph: Paragraph): Paragraph => {
		const name = added.get(paragraph.id);
		return name ? { ...paragraph, bookmarks: [...(paragraph.bookmarks ?? []), name] } : paragraph;
	};
	return {
		...model,
		blocks: model.blocks.map((block) =>
			block.type === 'paragraph'
				? mark(block)
				: {
						...block,
						rows: block.rows.map((row) =>
							row.map((cell) => ({ ...cell, paragraphs: cell.paragraphs.map(mark) })),
						),
					},
		),
	};
}

/**
 * The TOC as body paragraphs: one per entry, wrapped in a complex field whose begin, code and
 * separator open the first paragraph and whose end closes the last. Entries use Word's `TOCn`
 * styles when the document defines them and matching direct indents otherwise.
 */
export function buildTableOfContents(model: DocumentModel, options: TocOptions): Paragraph[] {
	const instruction = options.instruction ?? DEFAULT_TOC_INSTRUCTION;
	const field = { instr: instruction.trim() };
	const width = options.contentWidthTwips ?? 9360;
	const styles = model.paragraphStyles?.styles ?? {};
	const entries = tocEntries(model, instruction);
	const paragraphs: Paragraph[] = (entries.length ? entries : [undefined]).map((entry) => {
		const level = entry?.level ?? 1;
		const style = `TOC${level}`;
		const page = entry ? options.pageNumbers?.get(entry.blockId) : undefined;
		const bookmark = entry ? options.bookmarks?.get(entry.blockId) : undefined;
		const runs: TextRun[] = !entry
			? [{ text: EMPTY_TOC_TEXT, bold: true, field }]
			: bookmark
				? linkedEntry(entry.text, bookmark, page)
				: [{ text: page ? `${entry.text}\t${page}` : entry.text, field }];
		return {
			type: 'paragraph',
			id: options.newId(),
			runs,
			...(styles[style]
				? { style }
				: { ...(level > 1 ? { indentLeftTwips: (level - 1) * 220 } : {}), spacingAfterTwips: 100 }),
			tabStops: [{ posTwips: width, align: 'right', leader: 'dot' }],
		};
	});
	paragraphs[0].runs.unshift(
		{ text: '', fieldChar: 'begin' },
		{ text: '', fieldCode: instruction },
		{ text: '', fieldChar: 'separate' },
	);
	paragraphs.at(-1)!.runs.push({ text: '', fieldChar: 'end' });
	return paragraphs;
}

export interface TocLocation {
	/** Body block indexes of the paragraphs holding the field's begin and end markers. */
	start: number;
	end: number;
	instruction: string;
	/** Runs before the begin marker and after the end marker, kept when the TOC is rebuilt. */
	before: TextRun[];
	after: TextRun[];
}

/** The first TOC field in the body, following nested fields with a marker stack. */
export function findTableOfContents(blocks: Block[]): TocLocation | undefined {
	const stack: { code: string; block: number; run: number }[] = [];
	for (let index = 0; index < blocks.length; index++) {
		const block = blocks[index];
		if (block.type !== 'paragraph') continue;
		for (const [runIndex, run] of block.runs.entries()) {
			if (run.fieldChar === 'begin') stack.push({ code: '', block: index, run: runIndex });
			else if (run.fieldCode !== undefined && stack.length)
				stack[stack.length - 1].code += run.fieldCode;
			else if (run.fieldChar === 'end') {
				const frame = stack.pop();
				if (!frame || fieldName(frame.code) !== 'TOC') continue;
				const first = blocks[frame.block] as Paragraph;
				return {
					start: frame.block,
					end: index,
					instruction: frame.code,
					before: first.runs.slice(0, frame.run),
					after: block.runs.slice(runIndex + 1),
				};
			}
		}
	}
	return undefined;
}

/** Replaces the document's TOC field with freshly built entries; returns undefined when there is none. */
export function updateTableOfContents(
	model: DocumentModel,
	options: Omit<TocOptions, 'instruction'>,
): DocumentModel | undefined {
	const found = findTableOfContents(model.blocks);
	if (!found) return undefined;
	const linked = tocHyperlinks(found.instruction)
		? tocBookmarks(model, tocEntries(model, found.instruction))
		: undefined;
	const rebuilt = buildTableOfContents(model, {
		...options,
		instruction: found.instruction,
		...(linked ? { bookmarks: linked.bookmarks } : {}),
	});
	rebuilt[0].runs.unshift(...found.before);
	rebuilt.at(-1)!.runs.push(...found.after);
	const blocks = [...withTocBookmarks(model, linked?.added ?? new Map()).blocks];
	blocks.splice(found.start, found.end - found.start + 1, ...rebuilt);
	return { ...model, blocks };
}
