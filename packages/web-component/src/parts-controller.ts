import type { Block, HeaderFooterContent, Note } from '@christophervr/docx-core';
import type { EditorHost } from './editor-host';
import { buildFooterElement, buildHeaderElement } from './header-footer-view';
import { attachHeaderFooterEditing, type HeaderFooterSlotName } from './header-footer-editor';
import { buildNotesElement } from './notes-view';
import { attachNoteEditing } from './note-editor';
import { insertNote, type NoteKind } from './note-commands';

/** Headers, footers, footnotes and endnotes around the continuous editing surface. */
export class PartsController {
	private headerEl?: HTMLElement;
	private footerEl?: HTMLElement;
	private notesEl?: HTMLElement;

	constructor(private readonly host: EditorHost) {}

	/** Rebuilds the header/footer/note previews and wires in-place editing. */
	render(canvas: HTMLElement | undefined, paper: HTMLElement | undefined): void {
		if (!canvas || !paper) return;
		this.headerEl?.remove();
		this.footerEl?.remove();
		this.notesEl?.remove();
		const model = this.host.model();
		const locale = this.host.locale();
		this.headerEl = buildHeaderElement(model, locale) ?? undefined;
		this.footerEl = buildFooterElement(model, locale) ?? undefined;
		this.notesEl = buildNotesElement(model, locale) ?? undefined;
		const editable = () => this.host.canEditOutsideBody();
		for (const [element, kind] of [
			[this.headerEl, 'headers'],
			[this.footerEl, 'footers'],
		] as const)
			if (element)
				attachHeaderFooterEditing(element, {
					content: (slot) => this.host.model().sections?.[0]?.[kind]?.[slot],
					change: (slot, blocks) => this.updateHeaderFooter(kind, slot, blocks),
					editable,
				});
		if (this.notesEl)
			attachNoteEditing(this.notesEl, {
				note: (id) => {
					const current = this.host.model();
					return [...(current.footnotes ?? []), ...(current.endnotes ?? [])].find(
						(note) => note.id === id,
					);
				},
				change: (id, blocks) => this.updateNote(id, blocks),
				editable,
			});
		if (this.headerEl) canvas.insertBefore(this.headerEl, paper);
		if (this.footerEl) canvas.insertBefore(this.footerEl, paper.nextSibling);
		if (this.notesEl) canvas.insertBefore(this.notesEl, (this.footerEl ?? paper).nextSibling);
	}

	/** Inserts a footnote or endnote reference at the selection and opens the new note for typing. */
	insertNote(kind: NoteKind, canvas?: HTMLElement, paper?: HTMLElement): void {
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const { model, id } = insertNote(view, this.host.model(), kind);
		const key = kind === 'footnote' ? 'footnotes' : 'endnotes';
		this.host.setModel({ ...this.host.model(), [key]: model[key] });
		this.render(canvas, paper);
		const item = this.notesEl?.querySelector<HTMLElement>(
			`.dve-notes-${kind} li[data-docx-note-id="${id}"]`,
		);
		item?.scrollIntoView?.({ block: 'center' });
		item?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
		this.host.edited();
	}

	/** Replaces one footnote's or endnote's blocks. */
	private updateNote(id: string, blocks: Block[]): void {
		const replace = (notes: Note[] | undefined) =>
			notes?.map((note) => (note.id === id ? { ...note, blocks: structuredClone(blocks) } : note));
		const model = this.host.model();
		this.host.setModel({
			...model,
			footnotes: replace(model.footnotes),
			endnotes: replace(model.endnotes),
		});
		this.host.edited();
	}

	/** Applies header/footer edits to every section slot that shares the edited part. */
	private updateHeaderFooter(
		kind: 'headers' | 'footers',
		slot: HeaderFooterSlotName,
		blocks: Block[],
	): void {
		const model = this.host.model();
		const edited = model.sections?.[0]?.[kind]?.[slot];
		if (!edited) return;
		const sections = (model.sections ?? []).map((section) => {
			const slots = { ...section[kind] };
			for (const [name, content] of Object.entries(slots) as [
				HeaderFooterSlotName,
				HeaderFooterContent,
			][])
				if (content === edited || (edited.partName && content?.partName === edited.partName))
					slots[name] = { ...content, blocks: structuredClone(blocks) };
			return { ...section, [kind]: slots };
		});
		this.host.setModel({ ...model, sections });
		this.host.edited();
	}
}
