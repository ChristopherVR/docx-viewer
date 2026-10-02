import type { Block, HeaderFooterContent, Note } from '@christophervr/docx-core';
import { TextSelection, type Plugin } from 'prosemirror-state';
import { closeHistory, undo, redo } from 'prosemirror-history';
import { runStylesPlugin } from './run-styles';
import { paragraphStylesPlugin } from './paragraph-styles';
import type { EditorView } from 'prosemirror-view';
import type { EditorHost } from './editor-host';
import { buildFooterElement, buildHeaderElement } from './header-footer-view';
import {
	attachHeaderFooterEditing,
	type HeaderFooterSlotName,
	type InlineEditorOptions,
} from './header-footer-editor';
import { imageNodeView, type ImageMediaCache } from './image-media';
import { buildNotesElement } from './notes-view';
import { attachNoteEditing } from './note-editor';
import { insertNote, type NoteKind } from './note-commands';
import { sectionPartsJson, HEADER_FOOTER_INPUT, SectionPartsStep } from './header-footer-history';
import { effectiveHeaderFooter, withHeaderFooterLink } from './header-footer-link';
import type { HeaderFooterContext } from './header-footer-ribbon';
import { storyPreview, selectSectionStart } from './header-footer-navigation';
import { withBlankHeaderFooter, newHeaderFooterId } from './header-footer-commands';
import { sectionLayoutJson } from './section-layout';

export interface PartsControllerHost extends EditorHost {
	images(): ImageMediaCache;
	/** Plugins shared with in-place editors (e.g. Ctrl+K). */
	plugins(): Plugin[];
	/** Where focus may go without closing an in-place editor: the ribbon and dialogs. */
	keepOpenWithin(): (Element | undefined)[];
	sectionIndex?(): number;
	refreshControls?(): void;
}

/** Headers, footers, footnotes and endnotes around the continuous editing surface. */
export class PartsController {
	private headerEl?: HTMLElement | undefined;
	private footerEl?: HTMLElement | undefined;
	private notesEl?: HTMLElement | undefined;
	private active?: EditorView | undefined;
	private closeActive?: (() => void) | undefined;
	private headerContext?: { kind: 'headers' | 'footers'; slot: HeaderFooterSlotName } | undefined;
	private shownSection = 0;
	private renderingContext?: HeaderFooterContext | undefined;

	constructor(private readonly host: PartsControllerHost) {}

	/** The in-place header/footer/note editor currently open, which the ribbon then targets. */
	activeView(): EditorView | undefined {
		return this.active;
	}
	/** Page-number insertion follows the first/even/default story currently being edited. */
	activeHeaderFooterSlot(): HeaderFooterSlotName | undefined {
		return this.headerContext?.slot;
	}
	headerFooterContext(): HeaderFooterContext | undefined {
		return this.headerContext
			? { ...this.headerContext, index: this.shownSection }
			: this.renderingContext;
	}
	closeHeaderFooter(canvas?: HTMLElement, paper?: HTMLElement): void {
		this.closeActive?.();
		this.render(canvas, paper);
		this.host.view()?.focus();
	}
	navigate(
		target: 'header' | 'footer' | 'previous' | 'next',
		canvas?: HTMLElement,
		paper?: HTMLElement,
	): void {
		const context = this.headerContext;
		const view = this.host.view();
		if (!context || !view) return;
		const index = this.shownSection + (target === 'previous' ? -1 : target === 'next' ? 1 : 0);
		if (index < 0 || index >= (this.host.model().sections?.length ?? 0)) return;
		this.shownSection = index;
		this.headerContext = {
			...context,
			kind: target === 'header' ? 'headers' : target === 'footer' ? 'footers' : context.kind,
		};
		selectSectionStart(view, this.host.model(), index);
		this.render(canvas, paper, true);
	}
	toggleLink(): void {
		const context = this.headerContext;
		const view = this.host.view();
		if (!context || !view?.editable || !this.host.canEditOutsideBody()) return;
		const model = this.host.model();
		const linked = !model.sections?.[this.shownSection]?.[context.kind]?.[context.slot];
		const next = withHeaderFooterLink(
			model,
			this.shownSection,
			context.kind,
			context.slot,
			!linked,
		);
		if (next !== model)
			view.dispatch(
				closeHistory(view.state.tr).setDocAttribute(
					'sectionParts',
					sectionPartsJson(next.sections),
				),
			);
	}

	/** Sets picture sources in a read-only preview from the package media. */
	private readonly decorate = (preview: HTMLElement) => {
		for (const image of preview.querySelectorAll<HTMLImageElement>('img[data-docx-image]')) {
			const src = this.host
				.images()
				.urlFor(image.dataset.partName ?? '', image.dataset.contentType ?? '');
			if (src) image.src = src;
		}
	};

	private editorOptions(header = false): InlineEditorOptions {
		return {
			contextModel: () => this.host.model(),
			nodeViews: {
				image: imageNodeView(this.host.images(), { theme: () => this.host.model().theme }),
			},
			decorate: this.decorate,
			plugins: [
				...this.host.plugins(),
				runStylesPlugin(() => this.host.model()),
				paragraphStylesPlugin(() => this.host.model()),
			],
			keepOpenWithin: () => this.host.keepOpenWithin(),
			activate: (view) => {
				this.active = view;
				if (header) this.host.view()?.dispatch(closeHistory(this.host.view()!.state.tr));
			},
			deactivate: (view) => {
				if (this.active === view) {
					this.active = undefined;
					this.closeActive = undefined;
				}
				if (header) {
					this.headerContext = undefined;
					this.host.view()?.dispatch(closeHistory(this.host.view()!.state.tr));
					this.host.refreshControls?.();
				}
			},
			registerClose: (close) => {
				this.closeActive = close;
			},
			...(header ? { history: (key: 'undo' | 'redo') => this.runHistory(key) } : {}),
		};
	}

	/** Header/footer typing and insertion use the body history; notes retain their local stack. */
	runHistory(key: 'undo' | 'redo'): boolean {
		const view = this.host.view();
		return !!view?.editable && (key === 'undo' ? undo : redo)(view.state, view.dispatch);
	}
	usesBodyHistory(): boolean {
		return !!this.headerContext;
	}

	syncSection(canvas?: HTMLElement, paper?: HTMLElement): void {
		if (!this.active && this.shownSection !== (this.host.sectionIndex?.() ?? 0))
			this.render(canvas, paper);
	}

	/** Rebuilds the header/footer/note previews and wires in-place editing. */
	render(
		canvas: HTMLElement | undefined,
		paper: HTMLElement | undefined,
		keepEditor = false,
	): void {
		if (!canvas || !paper) return;
		const context = keepEditor ? this.headerContext : undefined;
		const selection = context ? this.active?.state.selection : undefined;
		if (!context) this.shownSection = this.host.sectionIndex?.() ?? 0;
		this.renderingContext = context ? { ...context, index: this.shownSection } : undefined;
		this.closeActive?.();
		this.headerEl?.remove();
		this.footerEl?.remove();
		this.notesEl?.remove();
		const model = storyPreview(this.host.model(), this.renderingContext);
		const locale = this.host.locale();
		this.headerEl = buildHeaderElement(model, locale, this.shownSection) ?? undefined;
		this.footerEl = buildFooterElement(model, locale, this.shownSection) ?? undefined;
		this.notesEl = buildNotesElement(model, locale) ?? undefined;
		const editable = () => this.host.canEditOutsideBody();
		for (const [element, kind] of [
			[this.headerEl, 'headers'],
			[this.footerEl, 'footers'],
		] as const)
			if (element)
				attachHeaderFooterEditing(element, {
					content: (slot) =>
						effectiveHeaderFooter(this.host.model(), this.shownSection, kind, slot) ??
						effectiveHeaderFooter(model, this.shownSection, kind, slot),
					change: (slot, blocks) => this.updateHeaderFooter(kind, slot, blocks),
					editable,
					editor: {
						...this.editorOptions(true),
						activate: (view) => {
							this.editorOptions(true).activate?.(view);
							const slot = view.dom.closest<HTMLElement>('[data-slot]')?.dataset
								.slot as HeaderFooterSlotName;
							this.headerContext = { kind, slot };
							this.host.refreshControls?.();
						},
					},
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
				editor: this.editorOptions(),
			});
		for (const element of [this.headerEl, this.footerEl, this.notesEl])
			if (element) this.decorate(element);
		if (this.headerEl) canvas.insertBefore(this.headerEl, paper);
		if (this.footerEl) canvas.insertBefore(this.footerEl, paper.nextSibling);
		if (this.notesEl) canvas.insertBefore(this.notesEl, (this.footerEl ?? paper).nextSibling);
		if (context) {
			const root = context.kind === 'headers' ? this.headerEl : this.footerEl;
			root
				?.querySelector<HTMLElement>(`[data-slot="${context.slot}"]`)
				?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
			if (this.active && selection) {
				const size = this.active.state.doc.content.size;
				this.active.dispatch(
					this.active.state.tr.setSelection(
						TextSelection.near(this.active.state.doc.resolve(Math.min(selection.from, size))),
					),
				);
			} else this.host.view()?.focus();
		}
		this.renderingContext = undefined;
		this.host.refreshControls?.();
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
		const replace = (notes: Note[]) =>
			notes.map((note) => (note.id === id ? { ...note, blocks: structuredClone(blocks) } : note));
		const model = this.host.model();
		this.host.setModel({
			...model,
			...(model.footnotes && { footnotes: replace(model.footnotes) }),
			...(model.endnotes && { endnotes: replace(model.endnotes) }),
		});
		this.host.edited();
	}

	/** Applies header/footer edits to every section slot that shares the edited part. */
	private updateHeaderFooter(
		kind: 'headers' | 'footers',
		slot: HeaderFooterSlotName,
		blocks: Block[],
	): void {
		let model = this.host.model();
		if (!effectiveHeaderFooter(model, this.shownSection, kind, slot))
			model = withBlankHeaderFooter(model, kind, newHeaderFooterId, this.shownSection, slot);
		const edited = effectiveHeaderFooter(model, this.shownSection, kind, slot)!;
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
		const view = this.host.view();
		if (!view?.editable || !this.host.canEditOutsideBody()) return;
		const tr = view.state.tr;
		const layout = sectionLayoutJson(sections);
		if (layout !== view.state.doc.attrs.sections) tr.setDocAttribute('sections', layout);
		if (Boolean(model.evenAndOddHeaders) !== view.state.doc.attrs.evenAndOddHeaders)
			tr.setDocAttribute('evenAndOddHeaders', Boolean(model.evenAndOddHeaders));
		view.dispatch(
			tr.step(new SectionPartsStep(sectionPartsJson(sections))).setMeta(HEADER_FOOTER_INPUT, true),
		);
	}
}
