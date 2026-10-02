import type { DocumentModel, Note } from 'docx-core';
import { formatNoteNumber, numberNotesInOrder } from 'docx-core';
import { renderBlocks } from './header-footer-view';
import { translateUiText } from './localization';

function noteList(
	notes: Note[],
	kind: 'footnote' | 'endnote',
	numFmt: string | undefined,
	blocks: DocumentModel['blocks'],
	model: DocumentModel,
): HTMLElement {
	const order = numberNotesInOrder(blocks, kind);
	const numbered = notes
		.map((note, index) => ({ note, number: order.get(note.id) ?? order.size + index + 1 }))
		.sort((a, b) => a.number - b.number);
	const list = document.createElement('ol');
	list.className = `dve-note-list dve-note-list-${kind}`;
	for (const { note, number } of numbered) {
		const item = document.createElement('li');
		item.dataset.docxNoteId = note.id;
		item.value = number;
		const marker = document.createElement('span');
		marker.className = 'dve-note-marker';
		marker.textContent = `${formatNoteNumber(number, numFmt)}. `;
		const body = document.createElement('div');
		body.className = 'dve-note-body';
		body.append(renderBlocks(note.blocks, model));
		item.append(marker, body);
		list.append(item);
	}
	return list;
}

function section(
	label: string,
	notes: Note[] | undefined,
	kind: 'footnote' | 'endnote',
	numFmt: string | undefined,
	blocks: DocumentModel['blocks'],
	locale: string,
	model: DocumentModel,
): HTMLElement | null {
	if (!notes?.length) return null;
	const root = document.createElement('section');
	root.className = `dve-notes dve-notes-${kind}`;
	root.setAttribute('contenteditable', 'false');
	root.dataset.editorLocale = locale;
	root.dataset.localeAriaLabel = label;
	root.setAttribute('aria-label', translateUiText(root, label));
	const heading = document.createElement('h2');
	heading.className = 'dve-notes-heading';
	heading.textContent = translateUiText(root, label);
	heading.dataset.localeAriaLabel = label;
	root.append(heading, noteList(notes, kind, numFmt, blocks, model));
	return root;
}

/** Footnotes then endnotes, read-only, rendered at the end of the continuous surface. */
export function buildNotesElement(model: DocumentModel, locale: string): HTMLElement | null {
	const footnotes = section(
		'Footnotes',
		model.footnotes,
		'footnote',
		model.footnoteNumFmt,
		model.blocks,
		locale,
		model,
	);
	const endnotes = section(
		'Endnotes',
		model.endnotes,
		'endnote',
		model.endnoteNumFmt,
		model.blocks,
		locale,
		model,
	);
	if (!footnotes && !endnotes) return null;
	const root = document.createElement('div');
	root.className = 'dve-notes-container';
	if (footnotes) root.append(footnotes);
	if (endnotes) root.append(endnotes);
	return root;
}
