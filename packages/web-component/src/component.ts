import { EditorState, Transaction, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history, undo, redo } from 'prosemirror-history';
import { baseKeymap, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import type { DocumentModel, Paragraph } from '@christophervr/docx-core';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { schema } from './schema';
import { createRibbon, type RibbonAction } from './ribbon';
import { syncFontControls, syncParagraphControls, syncFormatControls } from './ribbon-controls';
import {
	applyFont,
	applyPageStyles,
	clearFormatting,
	insertTable,
	updateParagraphs,
	updatePage,
} from './ribbon-commands';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import styleText from './style.css?inline';
import { applyHighlight, toggleVerticalAlign } from './inline-commands';
import { executeTableCommand, canExecuteTableCommand } from './table-commands';

const HTMLElementBase: typeof HTMLElement =
	typeof HTMLElement === 'undefined' ? (class {} as typeof HTMLElement) : HTMLElement;
const markCommands = {
	bold: toggleMark(schema.marks.bold),
	italic: toggleMark(schema.marks.italic),
	underline: toggleMark(schema.marks.underline),
	strike: toggleMark(schema.marks.strike),
};
const editableCommand =
	(command: Command): Command =>
	(state, dispatch, view) =>
		view?.editable === false ? false : command(state, dispatch, view);

export class DocxEditorElement extends HTMLElementBase {
	private model: DocumentModel = createDocument();
	private view?: EditorView;
	private loaded?: Awaited<ReturnType<typeof loadDocument>>;
	private loadGeneration = 0;
	private _readOnly = false;
	private toolbar?: HTMLElement;
	private paper?: HTMLElement;
	private zoom = 1;

	get documentModel() {
		return this.model;
	}
	set documentModel(value: DocumentModel | null) {
		this.loadGeneration++;
		this.loaded = undefined;
		this.model = value || createDocument();
		if (this.isConnected) this.renderDocument();
	}

	get readOnly() {
		return this._readOnly;
	}
	set readOnly(value: boolean) {
		this._readOnly = Boolean(value);
		if (this.view) this.view.setProps({ editable: () => !this._readOnly });
		this.refreshControls();
	}

	connectedCallback() {
		this.buildShell();
		this.renderDocument();
	}

	disconnectedCallback() {
		this.view?.destroy();
		this.view = undefined;
	}

	async load(input: Uint8Array | ArrayBuffer): Promise<void> {
		const generation = ++this.loadGeneration;
		try {
			const session = await loadDocument(input);
			if (generation !== this.loadGeneration) return;
			this.loaded = session;
			this.model = session.model;
			if (this.isConnected) this.renderDocument();
		} catch (cause) {
			const error = cause instanceof Error ? cause : new Error(String(cause));
			if (generation === this.loadGeneration) {
				this.dispatchEvent(
					new CustomEvent('document-error', { detail: error, bubbles: true, composed: true }),
				);
			}
			throw error;
		}
	}

	setLoadedDocument(session: Awaited<ReturnType<typeof loadDocument>>) {
		this.loadGeneration++;
		this.loaded = session;
		this.model = session.model;
		if (this.isConnected) this.renderDocument();
	}

	async save(): Promise<Uint8Array> {
		return this.loaded ? this.loaded.save(this.model) : saveDocx(this.model);
	}

	private buildShell() {
		if (this.toolbar) return;
		this.classList.add('dve-host');
		const root = this.attachShadow({ mode: 'open' });
		const style = document.createElement('style');
		style.textContent = styleText;
		const frame = document.createElement('section');
		frame.className = 'dve-frame';
		const toolbar = createRibbon();
		toolbar.addEventListener('ribbon-action', (event) =>
			this.handleRibbonAction((event as CustomEvent<RibbonAction>).detail),
		);
		const status = document.createElement('span');
		status.className = 'dve-status';
		status.textContent = 'Page 1';
		toolbar.append(status);
		const canvas = document.createElement('main');
		canvas.className = 'dve-canvas';
		const paper = document.createElement('div');
		paper.className = 'dve-paper';
		paper.setAttribute('aria-label', 'Document page');
		canvas.append(paper);
		frame.append(toolbar, canvas);
		root.append(style, frame);
		this.toolbar = toolbar;
		this.paper = paper;
		this.setAttribute('role', 'region');
		this.setAttribute('aria-label', 'Document editor');
	}

	private renderDocument() {
		if (!this.paper) return;
		this.view?.destroy();
		this.paper.replaceChildren();
		applyPageStyles(this.paper, this.model, this.zoom);
		const state = EditorState.create({
			doc: modelToDoc(this.model),
			plugins: [
				history(),
				keymap(
					Object.fromEntries(
						Object.entries({
							...baseKeymap,
							'Mod-z': undo,
							'Mod-y': redo,
							'Mod-Shift-z': redo,
							'Mod-b': markCommands.bold,
							'Mod-i': markCommands.italic,
							'Mod-u': markCommands.underline,
						}).map(([key, command]) => [key, editableCommand(command)]),
					),
				),
			],
		});
		this.view = new EditorView(this.paper, {
			state,
			editable: () => !this._readOnly,
			dispatchTransaction: (transaction: Transaction) => this.applyTransaction(transaction),
		});
		this.refreshControls();
	}

	private applyTransaction(transaction: Transaction) {
		if (!this.view) return;
		const applied = this.view.state.applyTransaction(transaction).state;
		const repaired = assignMissingParagraphIds(applied);
		this.view.updateState(repaired ? applied.apply(repaired) : applied);
		if (transaction.docChanged) {
			this.model = docToModel(this.view.state.doc, this.model);
			if (this.paper) applyPageStyles(this.paper, this.model, this.zoom);
			this.dispatchEvent(
				new CustomEvent('document-change', { detail: this.model, bubbles: true, composed: true }),
			);
		}
		this.refreshControls();
	}

	private command(
		command: (
			state: EditorState,
			dispatch?: (transaction: Transaction) => void,
			view?: EditorView,
		) => boolean,
	) {
		if (this.view && !this._readOnly) command(this.view.state, this.view.dispatch, this.view);
	}

	private setAlignment(align: Paragraph['align']) {
		if (!this.view || this._readOnly) return;
		const { from, to } = this.view.state.selection;
		const positions: number[] = [];
		this.view.state.doc.nodesBetween(from, to, (node, pos) => {
			if (node.type.name === 'paragraph') positions.push(pos);
		});
		if (!positions.length && this.view.state.selection.$from.parent.type.name === 'paragraph')
			positions.push(this.view.state.selection.$from.before());
		let transaction = this.view.state.tr;
		for (const pos of positions) {
			const node = transaction.doc.nodeAt(pos);
			if (node) transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, align });
		}
		if (transaction.docChanged) this.view.dispatch(transaction);
	}

	private handleRibbonAction(action: RibbonAction) {
		if (this._readOnly && action.type !== 'zoom') return;
		if (action.type === 'format') {
			if (action.key === 'superscript' || action.key === 'subscript') {
				if (this.view) toggleVerticalAlign(this.view, action.key);
			} else this.command(markCommands[action.key]);
		} else if (action.type === 'align') this.setAlignment(action.value);
		else if (action.type === 'history') this.command(action.key === 'undo' ? undo : redo);
		else if (action.type === 'font' && this.view) {
			if (action.key === 'highlight') applyHighlight(this.view, action.value);
			else applyFont(this.view, action.key, action.value);
		} else if (action.type === 'tableEdit' && this.view) executeTableCommand(this.view, action.key);
		else if (action.type === 'clear' && this.view) clearFormatting(this.view);
		else if (action.type === 'table' && this.view) insertTable(this.view);
		else if (action.type === 'page') this.setPage(action.key, action.value);
		else if (action.type === 'paragraph' && this.view)
			updateParagraphs(this.view, action.key, action.value);
		else if (action.type === 'zoom') {
			this.zoom = action.value / 100;
			if (this.paper) applyPageStyles(this.paper, this.model, this.zoom);
		}
		if (action.type !== 'zoom' && typeof document.execCommand === 'function') this.view?.focus();
	}

	private setPage(key: 'margin' | 'orientation', value: string) {
		if (this.view) updatePage(this.view, key, value);
	}

	private refreshControls() {
		this.toolbar
			?.querySelectorAll<HTMLButtonElement | HTMLSelectElement>(
				'.ribbon-group button, .ribbon-group select:not([aria-label="Zoom"])',
			)
			.forEach((control) => {
				control.disabled = this._readOnly;
			});
		if (!this.view) return;
		const { state } = this.view;
		if (this.toolbar) {
			syncFormatControls(this.toolbar, state);
			for (const button of this.toolbar.querySelectorAll<HTMLButtonElement>(
				'button[data-action]',
			)) {
				const action = JSON.parse(button.dataset.action!) as RibbonAction;
				if (action.type === 'tableEdit')
					button.disabled = this._readOnly || !canExecuteTableCommand(this.view, action.key);
			}
			syncFontControls(this.toolbar, state);
			syncParagraphControls(this.toolbar, state);
		}
		const content = state.doc.textBetween(0, state.doc.content.size, ' ').trim();
		const words = content ? content.split(/\s+/).length : 0;
		const status = this.toolbar?.parentElement?.querySelector('.dve-status');
		if (status) status.textContent = `Page 1 · ${words} words`;
	}
}

export function registerDocxEditor() {
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
