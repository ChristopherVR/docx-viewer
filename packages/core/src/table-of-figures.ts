// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Tables of figures are `TOC \c "Label"` fields whose entries are the captions numbered by a
// matching `SEQ Label` field.
import type { DocumentModel, Paragraph } from './model.js';

/** The caption label a TOC instruction collects (`\c "Figure"`), or undefined for a heading TOC. */
export function tocCaptionLabel(instruction: string): string | undefined {
	return /\\c\s*"([^"]+)"/.exec(instruction)?.[1];
}

/** The instruction Word writes for a table of figures of `label`. */
export const tableOfFiguresInstruction = (label: string) => ` TOC \\h \\z \\c "${label}" `;

/** Whether `paragraph` holds a `SEQ label` field, i.e. is a caption of that label. */
export function isCaptionOf(paragraph: Paragraph, label: string): boolean {
	const wanted = label.toLowerCase();
	return paragraph.runs.some((run) => {
		const instr = run.field?.instr ?? run.fieldCode;
		const match = instr ? /^\s*SEQ\s+("[^"]+"|\S+)/i.exec(instr) : null;
		return match?.[1]?.replace(/^"|"$/g, '').toLowerCase() === wanted;
	});
}

/** Caption paragraphs of `label` in body order (including those in tables). */
export function captionParagraphs(model: DocumentModel, label: string): Paragraph[] {
	return model.blocks
		.flatMap((block) =>
			block.type === 'paragraph'
				? [block]
				: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
		)
		.filter((paragraph) => isCaptionOf(paragraph, label));
}
