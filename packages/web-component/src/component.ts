import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { createRibbon, type RibbonAction } from './ribbon';
import { syncFontControls, syncParagraphControls, syncFormatControls } from './ribbon-controls';
import { applyPageStyles } from './ribbon-commands';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import styleText from './style.css?inline';
import { canExecuteTableCommand } from './table-commands';
import { editorKeymap, runRibbonCommand } from './editor-commands';
import { countWords } from './word-count';
import { createSearchPanel, type SearchPanelHandle } from './search-panel';
import {
	CollaborationClient,
	type CollaborationConfig,
	type ClientReceiveResult,
	type StepBatch,
} from './collaboration';
import {
	repairCollaborativeDocumentIds,
	createCollaborationIdGenerator,
} from './collaboration-identity';
import { syncMultilingualControls } from './multilingual-ribbon';

const HTMLElementBase: typeof HTMLElement =
	typeof HTMLElement === 'undefined' ? (class {} as typeof HTMLElement) : HTMLElement;
export class DocxEditorElement extends HTMLElementBase {
	private model: DocumentModel = createDocument();
	private view?: EditorView;
	private loaded?: Awaited<ReturnType<typeof loadDocument>>;
	private loadGeneration = 0;
	private _readOnly = false;
	private toolbar?: HTMLElement;
	private paper?: HTMLElement;
	private zoom = 1;
	private searchPanel?: SearchPanelHandle;
	private collaboration?: CollaborationClient;
	private collaborationIds?: (kind: string) => string;
	private detachedState?: EditorState;
	private sendScheduled = false;

	get documentModel() {
		return this.model;
	}
	set documentModel(value: DocumentModel | null) {
		this.assertDocumentReplaceable();
		this.detachedState = undefined;
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
		this.detachedState = this.view?.state;
		this.view?.destroy();
		this.view = undefined;
	}

	async load(input: Uint8Array | ArrayBuffer): Promise<void> {
		this.assertDocumentReplaceable();
		const generation = ++this.loadGeneration;
		try {
			const session = await loadDocument(input);
			if (generation !== this.loadGeneration) return;
			this.loaded = session;
			this.model = session.model;
			this.detachedState = undefined;
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
		this.assertDocumentReplaceable();
		this.detachedState = undefined;
		this.loadGeneration++;
		this.loaded = session;
		this.model = session.model;
		if (this.isConnected) this.renderDocument();
	}

	async save(): Promise<Uint8Array> {
		return this.loaded ? this.loaded.save(this.model) : saveDocx(this.model);
	}

	/** Join only after loading the authority's matching document snapshot and version. */
	startCollaboration(config: CollaborationConfig): void {
		if (this.collaboration) throw new Error('Stop the current collaboration session first.');
		if (!this.view) throw new Error('Mount and load the document before starting collaboration.');
		this.collaboration = new CollaborationClient(config);
		this.collaborationIds = createCollaborationIdGenerator(config.clientId);
		this.loadGeneration++;
		this.detachedState = undefined;
		this.renderDocument();
	}

	getPendingCollaboration(): StepBatch | null {
		const state = this.view?.state ?? this.detachedState;
		return state && this.collaboration ? this.collaboration.createPendingBatch(state) : null;
	}

	receiveCollaboration(batch: unknown): ClientReceiveResult['status'] {
		if (!this.collaboration || !this.view) throw new Error('No mounted collaboration session.');
		const result = this.collaboration.receive(this.view.state, batch);
		if (result.status === 'applied') this.applyTransaction(result.transaction, true);
		return result.status;
	}

	/** Stopping with pending edits requires an explicit discard of the transport queue. */
	stopCollaboration(discardPending = false): void {
		const state = this.view?.state ?? this.detachedState;
		if (!discardPending && state && this.collaboration?.pendingStepCount(state))
			throw new Error(
				'Acknowledge pending collaboration edits before stopping, or explicitly discard the queue.',
			);
		this.collaboration = undefined;
		this.collaborationIds = undefined;
		this.detachedState = undefined;
		if (this.isConnected) this.renderDocument();
	}

	private assertDocumentReplaceable() {
		if (this.collaboration) throw new Error('Stop collaboration before replacing the document.');
	}

	private scheduleCollaborationSend() {
		if (!this.collaboration || this.sendScheduled) return;
		this.sendScheduled = true;
		queueMicrotask(() => {
			this.sendScheduled = false;
			if (!this.isConnected) return;
			const batch = this.getPendingCollaboration();
			if (batch)
				this.dispatchEvent(
					new CustomEvent('collaboration-send', { detail: batch, bubbles: true, composed: true }),
				);
		});
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
		this.searchPanel = createSearchPanel({
			getView: () => this.view,
			onClose: () => this.view?.focus(),
		});
		frame.append(toolbar, this.searchPanel.element, canvas);
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
		const state =
			this.detachedState ??
			EditorState.create({
				doc: modelToDoc(this.model),
				plugins: [
					history(),
					editorKeymap(() => this.showSearch()),
					...(this.collaboration ? [this.collaboration.plugin] : []),
				],
			});
		this.view = new EditorView(this.paper, {
			state,
			editable: () => !this._readOnly,
			dispatchTransaction: (transaction: Transaction) => this.applyTransaction(transaction),
		});
		this.detachedState = undefined;
		this.refreshControls();
		this.scheduleCollaborationSend();
	}

	private applyTransaction(transaction: Transaction, remote = false) {
		if (!this.view) return;
		const applied = this.view.state.applyTransaction(transaction).state;
		const repaired = remote
			? null
			: this.collaboration
				? repairCollaborativeDocumentIds(
						applied,
						this.collaboration.clientId,
						this.collaborationIds,
					)
				: assignMissingParagraphIds(applied);
		this.view.updateState(repaired ? applied.apply(repaired) : applied);
		if (transaction.docChanged) {
			this.model = docToModel(this.view.state.doc, this.model);
			if (this.paper) applyPageStyles(this.paper, this.model, this.zoom);
			this.dispatchEvent(
				new CustomEvent('document-change', { detail: this.model, bubbles: true, composed: true }),
			);
		}
		this.refreshControls();
		if (transaction.docChanged || remote) this.scheduleCollaborationSend();
	}

	private showSearch() {
		this.searchPanel?.open();
	}

	private handleRibbonAction(action: RibbonAction) {
		if (action.type === 'search') this.showSearch();
		else if (action.type === 'zoom') {
			this.zoom = action.value / 100;
			if (this.paper) applyPageStyles(this.paper, this.model, this.zoom);
		} else if (this.view) {
			runRibbonCommand(this.view, action, this.collaborationIds);
			if (typeof document.execCommand === 'function') this.view.focus();
		}
	}

	private refreshControls() {
		this.searchPanel?.refresh();
		this.toolbar
			?.querySelectorAll<HTMLButtonElement | HTMLSelectElement | HTMLInputElement>(
				'.ribbon-group button, .ribbon-group input, .ribbon-group select:not([aria-label="Zoom"])',
			)
			.forEach((control) => {
				control.disabled =
					this._readOnly && control.getAttribute('aria-label') !== 'Find and replace';
			});
		if (!this.view) return;
		const { state } = this.view;
		if (this.toolbar) {
			syncMultilingualControls(this.toolbar, state);
			syncFormatControls(this.toolbar, state);
			for (const button of this.toolbar.querySelectorAll<HTMLButtonElement>(
				'button[data-action]',
			)) {
				const action = JSON.parse(button.dataset.action!) as RibbonAction;
				if (action.type === 'tableEdit')
					button.disabled =
						this._readOnly ||
						Boolean(this.collaboration) ||
						!canExecuteTableCommand(this.view, action.key);
			}
			syncFontControls(this.toolbar, state);
			syncParagraphControls(this.toolbar, state);
		}
		const content = state.doc.textBetween(0, state.doc.content.size, ' ').trim();
		const words = countWords(content, this.lang || undefined);
		const status = this.toolbar?.parentElement?.querySelector('.dve-status');
		if (status) status.textContent = `Page 1 · ${words} words`;
	}
}

export function registerDocxEditor() {
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
