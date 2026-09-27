// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reports precisely which imported features are unsupported, so parse.ts stays focused on
// building the document model. Never silently claims Word parity for anything listed here.
import type { Block, Paragraph } from './model.js';
import { hasSpecialBreak } from './breaks.js';
import { textContent, WORD_NS, type XmlDocument, type XmlElement } from './xml.js';

function hasAny(document: XmlDocument, names: string[]): boolean {
	return names.some((name) => document.getElementsByTagNameNS(WORD_NS, name).length > 0);
}

/** Warnings derived from raw XML feature scans (features not represented in the model at all). */
export function warningsFor(document: XmlDocument): string[] {
	const warnings: string[] = [];
	const features: [string[], string][] = [
		[['object'], 'Embedded OLE objects (w:object) are preserved but rendered as a placeholder.'],
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
			['numPr'],
			'List numbering is preserved as paragraph XML but is not represented in the document model.',
		],
		[
			['rStyle'],
			'Character style inheritance and theme font/color resolution are not modeled; displayed formatting may differ from Word.',
		],
	];
	for (const [names, message] of features) if (hasAny(document, names)) warnings.push(message);
	if (
		Array.from(document.getElementsByTagNameNS(WORD_NS, 'instrText')).some((el) =>
			/HYPERLINK/i.test(textContent(el)),
		)
	)
		warnings.push(
			'Complex HYPERLINK field codes (w:fldChar/w:instrText) are preserved as raw XML but not parsed as links; editing paragraphs containing them is rejected.',
		);
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

/** Visits every paragraph in document order, including those nested in table cells. */
export function forEachParagraph(blocks: Block[], visit: (paragraph: Paragraph) => void): void {
	for (const block of blocks) {
		if (block.type === 'paragraph') visit(block);
		else for (const row of block.rows) for (const cell of row) cell.paragraphs.forEach(visit);
	}
}

/** Warnings derived from the parsed model: image kinds and bookmarks the model can't fully represent. */
export function imageAndBookmarkWarnings(blocks: Block[]): string[] {
	const warnings: string[] = [];
	let hasPicture = false;
	let hasAnchored = false;
	const unsupportedKinds = new Set<string>();
	let hasBookmarks = false;
	forEachParagraph(blocks, (paragraph) => {
		if (paragraph.bookmarks?.length) hasBookmarks = true;
		for (const run of paragraph.runs) {
			if (!run.image) continue;
			if (run.image.unsupported) unsupportedKinds.add(run.image.unsupported);
			else hasPicture = true;
			if (run.image.anchored) hasAnchored = true;
		}
	});
	if (hasPicture)
		warnings.push(
			'Inline pictures render from their original bytes and can be deleted or newly inserted; crop, artistic effects, and non-inline text wrapping are not reproduced.',
		);
	if (hasAnchored)
		warnings.push(
			'Floating (anchored) picture position and text wrapping are not reproduced; they render as inline placeholders sized to the original picture.',
		);
	if (unsupportedKinds.size)
		warnings.push(
			`${Array.from(unsupportedKinds).join(', ')} content is rendered as a labeled placeholder and is not editable.`,
		);
	if (hasBookmarks)
		warnings.push(
			'Bookmark names are read-only; editing a paragraph that contains bookmarks preserves them but moves their start/end to the paragraph boundaries.',
		);
	return warnings;
}
