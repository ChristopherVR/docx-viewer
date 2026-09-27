import { refreshEditorControls } from './editor-controls';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { createRibbon, setRibbonLocale, type RibbonAction } from './ribbon';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import styleText from './style.css?inline';
import chromeStyleText from './chrome.css?inline';
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
import { runListAction } from './list-commands';
import { sectionBreaksPlugin } from './section-commands';
import { noteNumberingPlugin } from './note-commands';
import { keymap } from 'prosemirror-keymap';
import { createPrintLayoutController, type PrintLayoutController } from './print-layout-view';
import { moveCursorToBlock } from './print-layout-cursor';
import { trackChangesPlugin, REMOTE_TRANSACTION_META } from './track-changes-mode';
import { reviewDisplayPlugin, type ReviewDisplayMode } from './review-display';
import { ReviewController } from './review-controller';
import { ImageMediaCache, imageNodeView } from './image-media';
import { EditorChrome } from './editor-chrome';
import { dispatchDocumentError, type EditorHost } from './editor-host';
import { PartsController } from './parts-controller';
import { PageController } from './page-controller';
import { focusView } from './focus-view';
import { InsertController } from './insert-controller';
import { runStylesPlugin } from './run-styles';
import { countWords } from './word-count';

const HTMLElementBase: typeof HTMLElement =
	typeof HTMLElement === 'undefined' ? (class {} as typeof HTMLElement) : HTMLElement;
export class DocxEditorElement extends HTMLElementBase {
	private model: DocumentModel = createDocument();
	private view?: EditorView;
	private loaded?: Awaited<ReturnType<typeof loadDocument>>;
	private loadGeneration = 0;
	private _readOnly = false;
	private toolbar?: HTMLElement;
	private canvas?: HTMLElement;
	private paper?: HTMLElement;
	private searchPanel?: SearchPanelHandle;
	private collaboration?: CollaborationClient;
	private collaborationIds?: (kind: string) => string;
	private detachedState?: EditorState;
	private sendScheduled = false;
	private presence?: EditorPresence;
	private _locale: EditorLocale = 'en';
	private printLayout?: PrintLayoutController;
	private readonly inserts = new InsertController({
		view: () => this.view,
		model: () => this.model,
		contentWidth: () =>
			this.model.page.width - this.model.page.marginLeft - this.model.page.marginRight,
		paper: () => this.paper,
		toolbar: () => this.toolbar,
		reportError: (error) =>
			this.dispatchEvent(
				new CustomEvent('document-error', { detail: error, bubbles: true, composed: true }),
			),
	});
	private readonly imageMedia = new ImageMediaCache((partName) =>
		this.inserts.media(partName, this.loaded?.media),
	);
	private reviewDisplayMode: ReviewDisplayMode = 'all';
	private review?: ReviewController;
	private _reviewAuthor = 'Author';
	private chrome?: EditorChrome;
	private readonly host: EditorHost = {
		element: this,
		view: () => this.view,
		model: () => this.model,
		setModel: (model) => {
			this.model = model;
		},
		locale: () => this._locale,
		canEditOutsideBody: () => !this._readOnly && !this.collaboration,
		edited: () => this.markEditedOutsideBody(),
		reportError: (cause) => dispatchDocumentError(this, cause),
	};
	private readonly parts = new PartsController(this.host);
	private readonly pages = new PageController({
		...this.host,
		paper: () => this.paper,
		toolbar: () => this.toolbar,
		statusBar: () => this.chrome?.statusBar,
		printLayout: () => this.printLayout,
		refreshControls: () => this.refreshControls(),
	});
	private pendingFileName?: string;

	get reviewAuthor(): string {
		return this._reviewAuthor;
	}
	set reviewAuthor(value: string) {
		this._reviewAuthor = value || 'Author';
	}

	/** File name shown in the title bar and used for downloads from the built-in File commands. */
	get fileName(): string {
		return this.chrome?.fileName ?? this.pendingFileName ?? 'Document1.docx';
	}
	set fileName(value: string) {
		if (this.chrome) this.chrome.fileName = value;
		else this.pendingFileName = value;
	}

	get locale(): string {
		return this._locale;
	}
	set locale(value: string) {
		this._locale = normalizeEditorLocale(value);
		if (this.toolbar) setRibbonLocale(this.toolbar, this._locale);
		this.searchPanel?.setLocale(this._locale);
		this.parts.render(this.canvas, this.paper);
		this.review?.setLocale(this._locale);
		this.chrome?.setLocale(this._locale);
		this.inserts.setLocale(this._locale);
		this.refreshControls();
	}

	publishPresence(profile: { name: string; color: string }) {
		if (!this.presence) throw new Error('Start collaboration before publishing presence.');
		this._reviewAuthor = profile.name || this._reviewAuthor;
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
		this.inserts.reset();
		this.chrome?.setSaveState('saved');
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
		this.imageMedia.release();
	}

	async load(input: Uint8Array | ArrayBuffer): Promise<void> {
		this.assertDocumentReplaceable();
		const generation = ++this.loadGeneration;
		try {
			const session = await loadDocument(input);
			if (generation !== this.loadGeneration) return;
			this.imageMedia.release();
			this.inserts.reset();
			this.loaded = session;
			this.model = session.model;
			this.detachedState = undefined;
			this.chrome?.setSaveState('saved');
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
		this.inserts.reset();
		this.loaded = session;
		this.model = session.model;
		if (this.isConnected) this.renderDocument();
	}

	async save(): Promise<Uint8Array> {
		// Only pass staged pictures when there are any: legacy DOC sessions take the model alone.
		const media = this.inserts.pendingMedia.size ? this.inserts.pendingMedia : undefined;
		if (this.loaded)
			return media ? this.loaded.save(this.model, media) : this.loaded.save(this.model);
		return saveDocx(this.model, media);
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
		style.textContent = `${styleText}
${chromeStyleText}`;
		const frame = document.createElement('section');
		frame.className = 'dve-frame';
		const toolbar = createRibbon(this._locale);
		toolbar.addEventListener('ribbon-action', (event) =>
			this.handleRibbonAction((event as CustomEvent<RibbonAction>).detail),
		);
		const canvas = document.createElement('main');
		canvas.className = 'dve-canvas';
		const paper = document.createElement('div');
		paper.className = 'dve-paper';
		paper.setAttribute('aria-label', 'Document page');
		canvas.append(paper);
		this.printLayout = createPrintLayoutController(canvas, (blockId, offset) => {
			this.pages.setViewMode('draft');
			if (this.view) moveCursorToBlock(this.view, blockId, offset);
		});
		canvas.append(this.printLayout.element);
		canvas.addEventListener('scroll', () => {
			if (this.pages.viewMode === 'print') {
				this.printLayout?.refreshCurrentPage();
				this.refreshControls();
			}
		});
		this.searchPanel = createSearchPanel({
			getView: () => this.view,
			onClose: () => this.view?.focus(),
		});
		this.searchPanel.setLocale(this._locale);
		this.review = new ReviewController({
			getModel: () => this.model,
			setModel: (model) => {
				this.model = model;
			},
			getView: () => this.view,
			getReviewAuthor: () => this._reviewAuthor,
			getCollaborationIds: () => this.collaborationIds,
			notifyChange: () =>
				this.dispatchEvent(
					new CustomEvent('document-change', { detail: this.model, bubbles: true, composed: true }),
				),
			refresh: () => this.refreshControls(),
		});
		this.review.setLocale(this._locale);
		this.inserts.setLocale(this._locale);
		const body = document.createElement('div');
		body.className = 'dve-body';
		body.append(canvas, this.review.commentsPanel.element);
		frame.append(toolbar, this.searchPanel.element, body);
		this.chrome = this.createChrome();
		this.chrome.mount(frame, toolbar);
		frame.append(
			this.inserts.linkDialog.element,
			this.inserts.pictureDialog.element,
			this.inserts.pictureInput,
		);
		if (this.pendingFileName) this.chrome.fileName = this.pendingFileName;
		this.chrome.setLocale(this._locale);
		root.append(style, frame);
		this.toolbar = toolbar;
		this.canvas = canvas;
		this.paper = paper;
		this.setAttribute('role', 'region');
		this.setAttribute('aria-label', 'Document editor');
	}

	private renderDocument() {
		if (!this.paper) return;
		resetStylePicker(this.toolbar);
		this.view?.destroy();
		this.paper.replaceChildren();
		this.pages.refreshPageStyles();
		const state =
			this.detachedState ??
			EditorState.create({
				doc: modelToDoc(this.model),
				plugins: [
					history(),
					...this.inserts.plugins(),
					noteNumberingPlugin(),
					sectionBreaksPlugin(),
					keymap({
						'Mod-Alt-f': () => (this.parts.insertNote('footnote', this.canvas, this.paper), true),
						'Mod-Alt-d': () => (this.parts.insertNote('endnote', this.canvas, this.paper), true),
					}),
					runStylesPlugin(() => this.model),
					paragraphStylesPlugin(() => this.model),
					trackChangesPlugin(
						() => this._reviewAuthor,
						() => Boolean(this.model.trackChanges),
					),
					reviewDisplayPlugin(() => this.reviewDisplayMode),
					editorKeymap(() => this.showSearch()),
					...(this.collaboration ? [this.collaboration.plugin] : []),
					...(this.presence ? [this.presence.client.plugin] : []),
				],
			});
		this.view = new EditorView(this.paper, {
			state,
			editable: () => !this._readOnly,
			dispatchTransaction: (transaction: Transaction) => this.applyTransaction(transaction),
			nodeViews: {
				image: imageNodeView(this.imageMedia, {
					editPicture: (pos) => this.inserts.pictureDialog.open(pos),
					maxWidth: () =>
						this.model.page.width - this.model.page.marginLeft - this.model.page.marginRight,
				}),
			},
			handleClick: (view, pos, event) => this.inserts.handleClick(view, pos, event),
		});
		this.detachedState = undefined;
		this.inserts.syncPaper();
		this.parts.render(this.canvas, this.paper);
		this.pages.relayout();
		this.refreshControls();
		this.scheduleCollaborationSend();
	}

	private applyTransaction(transaction: Transaction, remote = false) {
		if (!this.view) return;
		if (remote) transaction.setMeta(REMOTE_TRANSACTION_META, true);
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
			this.pages.refreshPageStyles();
			this.pages.relayout();
			this.chrome?.setSaveState('dirty');
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
		if (this.inserts.handle(action)) return;
		if (action.type === 'search') this.showSearch();
		else if (action.type === 'zoom') this.pages.setZoom(action.value);
		else if (action.type === 'list' && this.view) {
			runListAction(this.view, action.key, this.model);
			focusView(this.view);
		} else if (action.type === 'view') this.pages.setViewMode(action.value);
		else if (action.type === 'print') this.pages.print();
		else if (action.type === 'insertNote')
			this.parts.insertNote(action.kind, this.canvas, this.paper);
		else if (action.type === 'page') this.pages.changePageSetup(action.key, action.value);
		else if (action.type === 'sectionBreak') this.pages.insertSectionBreak(action.kind);
		else if (action.type === 'reviewDisplay') {
			this.reviewDisplayMode = action.value;
			this.view?.dispatch(this.view.state.tr);
		} else if (action.type === 'review') this.review?.handleReview(action.key);
		else if (action.type === 'comments') this.review?.handleComments(action.key);
		else if (this.view) {
			runRibbonCommand(this.view, action, this.collaborationIds);
			if (typeof document.execCommand === 'function') this.view.focus();
		}
	}

	private createChrome(): EditorChrome {
		return new EditorChrome({
			element: this,
			model: () => this.model,
			ribbon: () => this.toolbar,
			wordCount: () => {
				const doc = this.view?.state.doc;
				return doc
					? countWords(doc.textBetween(0, doc.content.size, ' ').trim(), this.lang || undefined)
					: 0;
			},
			locale: () => this._locale,
			readOnly: () => this._readOnly,
			setReadOnly: (readOnly) => {
				this.readOnly = readOnly;
				this.dispatchEvent(
					new CustomEvent('readonly-change', { detail: readOnly, bubbles: true, composed: true }),
				);
			},
			newDocument: () => {
				this.documentModel = createDocument();
			},
			load: (bytes) => this.load(bytes),
			save: () => this.save(),
			print: () => this.pages.print(),
			history: (key) => {
				if (!this.view) return;
				runRibbonCommand(this.view, { type: 'history', key }, this.collaborationIds);
				if (typeof document.execCommand === 'function') this.view.focus();
			},
			toggleComments: () => this.review?.handleComments('toggle'),
			setViewMode: (mode) => this.pages.setViewMode(mode),
			setZoom: (percent) => this.pages.setZoom(percent),
			reportError: (error) => dispatchDocumentError(this, error),
		});
	}

	private markEditedOutsideBody() {
		this.chrome?.setSaveState('dirty');
		this.pages.relayout();
		this.dispatchEvent(
			new CustomEvent('document-change', { detail: this.model, bubbles: true, composed: true }),
		);
	}

	private refreshControls() {
		this.searchPanel?.refresh();
		if (this.pages.viewMode === 'print') this.printLayout?.refreshCurrentPage();
		const status = refreshEditorControls(
			this.toolbar,
			this.view,
			this.model,
			this._readOnly,
			Boolean(this.collaboration),
			this._locale,
			this.lang,
			this.printLayout?.pageStatus(),
			Boolean(this.review?.commentsOpen),
		);
		if (status) this.chrome?.refresh(status.pageText, status.wordText);
		this.pages.syncControls();
		this.chrome?.titleBar.setCommentsOpen(Boolean(this.review?.commentsOpen));
	}
}

export function registerDocxEditor() {
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
