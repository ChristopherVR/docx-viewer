import { refreshEditorControls } from './editor-controls';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { createRibbon, setRibbonLocale, type RibbonAction } from './ribbon';
import { applyPageStyles } from './ribbon-commands';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import styleText from './style.css?inline';
import { editorKeymap, runRibbonCommand } from './editor-commands';
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
import { normalizeEditorLocale, type EditorLocale } from './localization';
import { EditorPresence } from './editor-presence';
import { paragraphStylesPlugin, resetStylePicker } from './paragraph-styles';

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
	private presence?: EditorPresence;
	private _locale: EditorLocale = 'en';

	get locale(): string {
		return this._locale;
	}
	set locale(value: string) {
		this._locale = normalizeEditorLocale(value);
		if (this.toolbar) setRibbonLocale(this.toolbar, this._locale);
		this.searchPanel?.setLocale(this._locale);
		this.refreshControls();
	}

	publishPresence(profile: { name: string; color: string }) {
		if (!this.presence) throw new Error('Start collaboration before publishing presence.');
		return this.presence.publish(profile);
	}
	receivePresence(message: unknown) {
		if (!this.presence) throw new Error('Start collaboration before receiving presence.');
		return this.presence.receive(message);
	}
	leavePresence() {
		return this.presence?.leave() ?? null;
	}

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
		this.presence?.leave();
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
		this.presence = new EditorPresence(config, this, () => this.view);
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
		this.presence?.leave();
		this.presence = undefined;
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
		const toolbar = createRibbon(this._locale);
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
		this.searchPanel.setLocale(this._locale);
		frame.append(toolbar, this.searchPanel.element, canvas);
		root.append(style, frame);
		this.toolbar = toolbar;
		this.paper = paper;
		this.setAttribute('role', 'region');
		this.setAttribute('aria-label', 'Document editor');
	}

	private renderDocument() {
		if (!this.paper) return;
		resetStylePicker(this.toolbar);
		this.view?.destroy();
		this.paper.replaceChildren();
		applyPageStyles(this.paper, this.model, this.zoom);
		const state =
			this.detachedState ??
			EditorState.create({
				doc: modelToDoc(this.model),
				plugins: [
					history(),
					paragraphStylesPlugin(() => this.model),
					editorKeymap(() => this.showSearch()),
					...(this.collaboration ? [this.collaboration.plugin] : []),
					...(this.presence ? [this.presence.client.plugin] : []),
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
		if (transaction.docChanged || transaction.selectionSet || remote) this.presence?.schedule();
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
		refreshEditorControls(
			this.toolbar,
			this.view,
			this.model,
			this._readOnly,
			Boolean(this.collaboration),
			this._locale,
			this.lang,
		);
	}
}

export function registerDocxEditor() {
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
