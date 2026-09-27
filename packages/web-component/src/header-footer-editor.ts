import { EditorState, type Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import { createDocument, type Block, type HeaderFooterContent } from '@christophervr/docx-core';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import { editorKeymap } from './editor-commands';
import { focusView } from './focus-view';
import { renderBlocks } from './header-footer-view';

export type HeaderFooterSlotName = 'default' | 'first' | 'even';

export interface HeaderFooterEditingOptions {
	/** The content currently shown in `slot`, if any. */
	content(slot: HeaderFooterSlotName): HeaderFooterContent | undefined;
	/** Called with the slot's new blocks after every edit. */
	change(slot: HeaderFooterSlotName, blocks: Block[]): void;
	/** Header/footer editing is off in Viewing mode and during collaboration. */
	editable(): boolean;
}

/**
 * Word-style header/footer editing on the continuous surface: double-click a header or footer
 * slot to edit it in place; Escape or moving focus elsewhere returns to the read-only preview.
 */
export function attachHeaderFooterEditing(
	root: HTMLElement,
	options: HeaderFooterEditingOptions,
): void {
	for (const slot of root.querySelectorAll<HTMLElement>('.dve-header-footer-slot')) {
		const name = slot.dataset.slot as HeaderFooterSlotName | undefined;
		const body = slot.querySelector<HTMLElement>('.dve-header-footer-body');
		if (!name || !body) continue;
		slot.title = 'Double-click to edit';
		slot.addEventListener('dblclick', () => {
			if (!options.editable() || slot.classList.contains('dve-header-footer-editing')) return;
			const content = options.content(name);
			if (content) openEditor(slot, body, name, content, options);
		});
	}
}

/**
 * Replaces `body` with an in-place editor for `blocks` (same schema, history and formatting keys as
 * the main editor). Escape or moving focus elsewhere closes it and re-renders a read-only preview.
 */
export function openBlocksEditor(
	container: HTMLElement,
	body: HTMLElement,
	blocks: Block[],
	change: (blocks: Block[]) => void,
): void {
	let model = { ...createDocument(), blocks: structuredClone(blocks) };
	const host = document.createElement('div');
	host.className = 'dve-header-footer-editor';
	body.replaceChildren(host);
	container.classList.add('dve-header-footer-editing');
	const view: EditorView = new EditorView(host, {
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [history(), editorKeymap(() => undefined)],
		}),
		dispatchTransaction(transaction: Transaction) {
			const applied = view.state.apply(transaction);
			const repaired = assignMissingParagraphIds(applied);
			view.updateState(repaired ? applied.apply(repaired) : applied);
			if (!transaction.docChanged) return;
			model = docToModel(view.state.doc, model);
			change(structuredClone(model.blocks));
		},
	});
	const close = () => {
		if (!container.classList.contains('dve-header-footer-editing')) return;
		container.classList.remove('dve-header-footer-editing');
		view.destroy();
		body.replaceChildren(renderBlocks(model.blocks));
	};
	host.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
	});
	host.addEventListener('focusout', (event) => {
		if (!host.contains(event.relatedTarget as Node | null)) close();
	});
	focusView(view);
}

function openEditor(
	slot: HTMLElement,
	body: HTMLElement,
	name: HeaderFooterSlotName,
	content: HeaderFooterContent,
	options: HeaderFooterEditingOptions,
): void {
	openBlocksEditor(slot, body, content.blocks, (blocks) => options.change(name, blocks));
}
