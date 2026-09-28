import { refreshEditorControls } from './editor-controls';
import { EditorState, Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { createRibbon, setRibbonLocale, type RibbonAction } from './ribbon';
import { assignMissingParagraphIds, docToModel, modelToDoc } from './model-adapter';
import styleText from './style.css?inline';
import chromeStyleText from './chrome.css?inline';
import { runRibbonCommand } from './editor-commands';
import { createSearchPanel, type SearchPanelHandle } from './search-panel';
import {
	type CollaborationConfig,
	type ClientReceiveResult,
	type StepBatch,
} from './collaboration';
import { repairCollaborativeDocumentIds } from './collaboration-identity';
import { normalizeEditorLocale, type EditorLocale } from './localization';
import { CollaborationSession } from './collaboration-session';
import { resetStylePicker } from './paragraph-styles';
import { bodyPlugins } from './editor-plugins';
import { runListAction } from './list-commands';
import { createPrintLayoutController, type PrintLayoutController } from './print-layout-view';
import { moveCursorToBlock } from './print-layout-cursor';
import { REMOTE_TRANSACTION_META } from './track-changes-mode';
import type { ReviewDisplayMode } from './review-display';
import { ReviewController } from './review-controller';
import { ImageMediaCache, imageNodeView } from './image-media';
import { EditorChrome } from './editor-chrome';
import { dispatchDocumentError, type EditorHost } from './editor-host';
import { PartsController } from './parts-controller';
import { PageController } from './page-controller';
import { focusView } from './focus-view';
import { InsertController } from './insert-controller';
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
	private readonly collab = new CollaborationSession(this, () => this.view);
	private detachedState?: EditorState;
	private _locale: EditorLocale = 'en';
	private printLayout?: PrintLayoutController;
	private readonly inserts = new InsertController({
		view: () => this.targetView(),
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
		canEditOutsideBody: () => !this._readOnly && !this.collab.client,
		edited: () => this.markEditedOutsideBody(),
		reportError: (cause) => dispatchDocumentError(this, cause),
	};
	private readonly parts = new PartsController({
		...this.host,
		images: () => this.imageMedia,
		plugins: () => this.inserts.plugins(),
		keepOpenWithin: () => [
			this.toolbar,
			this.inserts.linkDialog.element,
			this.inserts.pictureDialog.element,
		],
	});
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
		if (!this.collab.presence) throw new Error('Start collaboration before publishing presence.');
		this._reviewAuthor = profile.name || this._reviewAuthor;
		return this.collab.presence.publish(profile);
	}
	receivePresence(message: unknown) {
		if (!this.collab.presence) throw new Error('Start collaboration before receiving presence.');
		return this.collab.presence.receive(message);
	}
	leavePresence() {
		return this.collab.presence?.leave() ?? null;
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
		this.collab.presence?.leave();
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
		if (this.collab.active) throw new Error('Stop the current collaboration session first.');
		if (!this.view) throw new Error('Mount and load the document before starting collaboration.');
		this.collab.start(config);
		this.loadGeneration++;
		this.detachedState = undefined;
		this.renderDocument();
	}

	getPendingCollaboration(): StepBatch | null {
		const state = this.view?.state ?? this.detachedState;
		return state && this.collab.client ? this.collab.client.createPendingBatch(state) : null;
	}

	receiveCollaboration(batch: unknown): ClientReceiveResult['status'] {
		if (!this.collab.client || !this.view) throw new Error('No mounted collaboration session.');
		const result = this.collab.client.receive(this.view.state, batch);
		if (result.status === 'applied') this.applyTransaction(result.transaction, true);
		return result.status;
	}

	/** Stopping with pending edits requires an explicit discard of the transport queue. */
	stopCollaboration(discardPending = false): void {
		const state = this.view?.state ?? this.detachedState;
		if (!discardPending && state && this.collab.client?.pendingStepCount(state))
			throw new Error(
				'Acknowledge pending collaboration edits before stopping, or explicitly discard the queue.',
			);
		this.collab.stop();
		this.detachedState = undefined;
		if (this.isConnected) this.renderDocument();
	}

	private assertDocumentReplaceable() {
		if (this.collab.client) throw new Error('Stop collaboration before replacing the document.');
	}

	private scheduleCollaborationSend() {
		this.collab.scheduleSend(() => this.getPendingCollaboration());
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
		this.printLayout = createPrintLayoutController(
			canvas,
			(blockId, offset) => {
				this.pages.setViewMode('draft');
				if (this.view) moveCursorToBlock(this.view, blockId, offset);
			},
			(partName, contentType) => this.imageMedia.urlFor(partName, contentType),
		);
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
			getCollaborationIds: () => this.collab.ids,
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
				plugins: bodyPlugins({
					model: () => this.model,
					reviewAuthor: () => this._reviewAuthor,
					reviewDisplayMode: () => this.reviewDisplayMode,
					insertNote: (kind) => this.parts.insertNote(kind, this.canvas, this.paper),
					showSearch: () => this.showSearch(),
					extraPlugins: this.inserts.plugins(),
					collaborationPlugins: [
						...(this.collab.client ? [this.collab.client.plugin] : []),
						...(this.collab.presence ? [this.collab.presence.client.plugin] : []),
					],
				}),
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
			: this.collab.client
				? repairCollaborativeDocumentIds(applied, this.collab.client.clientId, this.collab.ids)
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
		if (transaction.docChanged || transaction.selectionSet || remote)
			this.collab.presence?.schedule();
	}

	private showSearch() {
		this.searchPanel?.open();
	}

	private handleRibbonAction(action: RibbonAction) {
		if (this.inserts.handle(action)) return;
		if (action.type === 'search') this.showSearch();
		else if (action.type === 'zoom') this.pages.setZoom(action.value);
		else if (action.type === 'list' && this.targetView()) {
			runListAction(this.targetView()!, action.key, this.model);
			focusView(this.targetView());
		} else if (action.type === 'view') this.pages.setViewMode(action.value);
		else if (action.type === 'print') this.pages.print();
		else if (action.type === 'insertNote')
			this.parts.insertNote(action.kind, this.canvas, this.paper);
		else if (action.type === 'page') this.pages.changePageSetup(action.key, action.value);
		else if (action.type === 'sectionBreak') this.pages.insertSectionBreak(action.kind);
		else if (action.type === 'evenOddHeaders') this.pages.toggleEvenOddHeaders();
		else if (action.type === 'reviewDisplay') {
			this.reviewDisplayMode = action.value;
			this.view?.dispatch(this.view.state.tr);
		} else if (action.type === 'review') this.review?.handleReview(action.key);
		else if (action.type === 'comments') this.review?.handleComments(action.key);
		else {
			const target = this.targetView();
			if (!target) return;
			runRibbonCommand(target, action, this.collab.ids);
			focusView(target);
		}
	}

	/** The ribbon acts on an open header/footer/note editor, otherwise on the main text. */
	private targetView(): EditorView | undefined {
		return this.parts.activeView() ?? this.view;
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
				runRibbonCommand(this.view, { type: 'history', key }, this.collab.ids);
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
			Boolean(this.collab.client),
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
