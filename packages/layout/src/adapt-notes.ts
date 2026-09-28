import {
	formatNoteNumber,
	numberNotesInOrder,
	type DocumentModel,
	type Note,
	type Paragraph,
} from '@christophervr/docx-core';
import type { LayoutNote, LayoutParagraph } from './input.js';

type NoteKind = 'footnote' | 'endnote';

/** Note numbers as Word shows them: in reference order, in each kind's number format. */
export function noteLabels(model: DocumentModel): (kind: NoteKind, id: string) => string {
	const order = {
		footnote: numberNotesInOrder(model.blocks, 'footnote'),
		endnote: numberNotesInOrder(model.blocks, 'endnote'),
	};
	return (kind, id) => {
		const number = order[kind].get(id);
		if (number === undefined) return '';
		return formatNoteNumber(
			number,
			kind === 'footnote'
				? (model.footnoteNumFmt ?? 'decimal')
				: (model.endnoteNumFmt ?? 'lowerRoman'),
		);
	};
}

/** Adapts a note's paragraphs, whose number mark (`w:footnoteRef`) shows the note's label. */
export function adaptNote(
	note: Note,
	label: string,
	adaptParagraph: (paragraph: Paragraph, markLabel?: string) => LayoutParagraph,
): LayoutParagraph[] {
	return note.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [adaptParagraph(block, label)]
			: block.rows.flatMap((row) =>
					row.flatMap((cell) =>
						cell.paragraphs.map((paragraph) => adaptParagraph(paragraph, label)),
					),
				),
	);
}

/** The footnotes a paragraph references, laid out at the bottom of the page it starts on. */
export function paragraphFootnotes(
	paragraph: Paragraph,
	model: DocumentModel,
	label: (kind: NoteKind, id: string) => string,
	adaptParagraph: (paragraph: Paragraph, markLabel?: string) => LayoutParagraph,
): LayoutNote[] {
	const notes: LayoutNote[] = [];
	for (const run of paragraph.runs) {
		const reference = run.noteReference;
		if (reference?.kind !== 'footnote') continue;
		const note = model.footnotes?.find((item) => item.id === reference.id);
		if (note)
			notes.push({
				id: note.id,
				paragraphs: adaptNote(note, label('footnote', note.id), adaptParagraph),
			});
	}
	return notes;
}

/** Endnotes in reference order, placed after the document's last paragraph as Word does. */
export function endnoteParagraphs(
	model: DocumentModel,
	label: (kind: NoteKind, id: string) => string,
	adaptParagraph: (paragraph: Paragraph, markLabel?: string) => LayoutParagraph,
): LayoutParagraph[] {
	const order = numberNotesInOrder(model.blocks, 'endnote');
	return [...order.keys()].flatMap((id) => {
		const note = model.endnotes?.find((item) => item.id === id);
		return note ? adaptNote(note, label('endnote', id), adaptParagraph) : [];
	});
}
