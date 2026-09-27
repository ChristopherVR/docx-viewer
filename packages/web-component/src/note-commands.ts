import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	formatNoteNumber,
	type DocumentModel,
	type Note,
	type Paragraph,
	type TextRun,
} from '@christophervr/docx-core';
import { schema } from './schema';

export type NoteKind = 'footnote' | 'endnote';

const STYLES: Record<NoteKind, { text: string; reference: string }> = {
	footnote: { text: 'FootnoteText', reference: 'FootnoteReference' },
	endnote: { text: 'EndnoteText', reference: 'EndnoteReference' },
};

/** The next free note id: Word numbers user notes from 1 (separators use -1 and 0). */
export function nextNoteId(notes: Note[] | undefined): string {
	const used = (notes ?? []).map((note) => Number(note.id)).filter(Number.isFinite);
	return String(Math.max(0, ...used) + 1);
}

/** The formatting Word gives a note reference and a note's own number mark. */
function referenceFormatting(
	model: DocumentModel,
	kind: NoteKind,
): Pick<TextRun, 'verticalAlign' | 'style'> {
	const style = STYLES[kind].reference;
	return {
		verticalAlign: 'superscript',
		...(model.characterStyles?.styles[style] ? { style } : {}),
	};
}

/**
 * Word's Insert Footnote/Endnote: adds a note (number mark plus a space, in the note text style
 * when the document defines one) and inserts its superscript reference at the selection.
 * Returns the updated model and the new note's id.
 */
export function insertNote(
	view: EditorView,
	model: DocumentModel,
	kind: NoteKind,
): { model: DocumentModel; id: string } {
	const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
	const id = nextNoteId(model[key]);
	const formatting = referenceFormatting(model, kind);
	const textStyle = STYLES[kind].text;
	const paragraph: Paragraph = {
		type: 'paragraph',
		id: `${kind === 'footnote' ? 'fn' : 'en'}${id}-p0`,
		runs: [{ text: '', noteMark: kind, ...formatting }, { text: ' ' }],
		...(model.paragraphStyles?.styles[textStyle] ? { style: textStyle } : {}),
	};
	const next: DocumentModel = {
		...model,
		[key]: [...(model[key] ?? []), { id, blocks: [paragraph] }],
	};
	const reference = schema.nodes.noteReference.create({
		kind,
		id,
		number: 1,
		format: JSON.stringify(formatting),
	});
	view.dispatch(view.state.tr.replaceSelectionWith(reference, false).scrollIntoView());
	return { model: next, id };
}

/** Keeps note reference numbers in document order, as Word renumbers notes after every change. */
export function noteNumberingPlugin(): Plugin {
	return new Plugin({
		appendTransaction(transactions, _oldState, state) {
			if (!transactions.some((transaction) => transaction.docChanged)) return null;
			const counters: Record<string, number> = { footnote: 0, endnote: 0 };
			const formats: Record<NoteKind, string> = {
				footnote: state.doc.attrs.footnoteNumFmt ?? 'decimal',
				endnote: state.doc.attrs.endnoteNumFmt ?? 'lowerRoman',
			};
			const tr = state.tr;
			state.doc.descendants((node, pos) => {
				if (node.type !== schema.nodes.noteReference) return true;
				const kind = node.attrs.kind as NoteKind;
				const number = ++counters[kind];
				const label = formatNoteNumber(number, formats[kind]);
				if (node.attrs.number !== number || node.attrs.label !== label)
					tr.setNodeMarkup(pos, undefined, { ...node.attrs, number, label });
				return false;
			});
			return tr.docChanged ? tr : null;
		},
	});
}
