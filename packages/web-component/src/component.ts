import { applyShellLabels } from './shell-labels';
import { EditorState } from 'prosemirror-state';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { setRibbonLocale } from './ribbon';
import type { CollaborationConfig, ClientReceiveResult, StepBatch } from './collaboration';
import { normalizeEditorLocale } from './localization';
import {
	applyThemeColors,
	normalizeThemeMode,
	type EditorTheme,
	type EditorThemeMode,
} from './theme';
import { EditorCore, type LoadedDocument } from './editor-core';
import { buildShell, type ShellApi } from './editor-shell';
import { renderDocument } from './editor-render';
import { DocxEditorApi } from './element-api';
import { emit, type DocxEditorEventMap } from './events';
import {
	DEFAULT_FILE_NAME,
	DEFAULT_REVIEW_AUTHOR,
	DOCX_EDITOR_ATTRIBUTES,
	applyAttribute,
	isDocxEditorAttribute,
	reflectAttribute,
} from './editor-attributes';

/**
 * Typed event overloads. Redeclaring the methods hides the inherited ones, so the standard DOM
 * event map and the string fallback are repeated here.
 */
export interface DocxEditorElement {
	addEventListener<K extends keyof DocxEditorEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: DocxEditorEventMap[K]) => unknown,
		options?: boolean | AddEventListenerOptions,
	): void;
	addEventListener<K extends keyof HTMLElementEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: HTMLElementEventMap[K]) => unknown,
		options?: boolean | AddEventListenerOptions,
	): void;
	addEventListener(
		type: string,
		listener: EventListenerOrEventListenerObject,
		options?: boolean | AddEventListenerOptions,
	): void;
	removeEventListener<K extends keyof DocxEditorEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: DocxEditorEventMap[K]) => unknown,
		options?: boolean | EventListenerOptions,
	): void;
	removeEventListener<K extends keyof HTMLElementEventMap>(
		type: K,
		listener: (this: DocxEditorElement, event: HTMLElementEventMap[K]) => unknown,
		options?: boolean | EventListenerOptions,
	): void;
	removeEventListener(
		type: string,
		listener: EventListenerOrEventListenerObject,
		options?: boolean | EventListenerOptions,
	): void;
}

export class DocxEditorElement extends DocxEditorApi {
	/** Attributes mirrored to the `locale`, `readOnly`, `fileName`, `reviewAuthor` and `theme` properties. */
	static get observedAttributes(): string[] {
		return [...DOCX_EDITOR_ATTRIBUTES];
	}

	protected readonly core = new EditorCore(this);
	private readonly shellApi: ShellApi = {
		setReadOnly: (readOnly) => {
			this.readOnly = readOnly;
		},
		setDocumentModel: (model) => {
			this.documentModel = model;
		},
		load: (bytes) => this.load(bytes),
		saveBytes: () => this.saveBytes(),
	};

	/** Internal seams that in-repo tests reach for; not public API. */
	private get view() {
		return this.core.view;
	}
	private get inserts() {
		return this.core.inserts;
	}

	attributeChangedCallback(name: string, _old: string | null, value: string | null) {
		if (isDocxEditorAttribute(name)) applyAttribute(this, name, value);
	}

	get reviewAuthor(): string {
		return this.core.reviewAuthor;
	}
	set reviewAuthor(value: string) {
		this.core.reviewAuthor = value || DEFAULT_REVIEW_AUTHOR;
		reflectAttribute(this, 'review-author', this.core.reviewAuthor);
	}

	/** File name shown in the title bar and used for downloads from the built-in File commands. */
	get fileName(): string {
		const { chrome } = this.core.shell;
		return chrome?.fileName ?? this.core.pendingFileName ?? DEFAULT_FILE_NAME;
	}
	set fileName(value: string) {
		const { chrome } = this.core.shell;
		if (chrome) chrome.fileName = value;
		else this.core.pendingFileName = value;
		reflectAttribute(this, 'file-name', this.fileName);
	}

	/** `light`, `dark`, or `auto` (default, follows the OS). Reflected to the `theme` attribute. */
	get theme(): EditorThemeMode {
		return this.core.theme;
	}
	set theme(value: EditorThemeMode) {
		this.core.theme = normalizeThemeMode(value);
		reflectAttribute(this, 'theme', this.core.theme);
	}
	/** Token overrides applied as inline `--dve-*` custom properties on the host. */
	get themeColors(): Partial<EditorTheme> | undefined {
		return this.core.themeColors;
	}
	set themeColors(value: Partial<EditorTheme> | undefined) {
		this.core.themeColors = value;
		this.core.appliedThemeVars = applyThemeColors(this, this.core.appliedThemeVars, value);
	}

	get locale(): string {
		return this.core.locale;
	}
	set locale(value: string) {
		const { core } = this;
		const { toolbar, searchPanel, review, chrome, canvas, paper } = core.shell;
		core.locale = normalizeEditorLocale(value);
		if (toolbar) setRibbonLocale(toolbar, core.locale);
		searchPanel?.setLocale(core.locale);
		core.parts.render(canvas, paper);
		review?.setLocale(core.locale);
		chrome?.setLocale(core.locale);
		core.inserts.setLocale(core.locale);
		core.formatDialogs.setLocale(core.locale);
		applyShellLabels(this, paper, core.locale);
		core.refreshControls();
		core.collab.presence?.relocalize();
		reflectAttribute(this, 'locale', core.locale);
	}

	publishPresence(profile: { name: string; color: string }) {
		if (!this.core.collab.presence)
			throw new Error('Start collaboration before publishing presence.');
		this.reviewAuthor = profile.name || this.reviewAuthor;
		return this.core.collab.presence.publish(profile);
	}
	receivePresence(message: unknown) {
		if (!this.core.collab.presence)
			throw new Error('Start collaboration before receiving presence.');
		return this.core.collab.presence.receive(message);
	}
	leavePresence() {
		return this.core.collab.presence?.leave() ?? null;
	}

	get documentModel() {
		return this.core.model;
	}
	set documentModel(value: DocumentModel | null) {
		this.assertDocumentReplaceable();
		this.replaceDocument(value || createDocument());
		this.core.shell.chrome?.setSaveState('saved');
		if (this.isConnected) renderDocument(this.core);
	}

	get readOnly() {
		return this.core.readOnly;
	}
	set readOnly(value: boolean) {
		this.core.readOnly = Boolean(value);
		this.core.view?.setProps({ editable: () => !this.core.readOnly });
		if (this.core.readOnly)
			this.core.parts.closeHeaderFooter(this.core.shell.canvas, this.core.shell.paper);
		this.core.refreshControls();
		reflectAttribute(this, 'read-only', this.core.readOnly);
	}

	connectedCallback() {
		buildShell(this.core, this.shellApi);
		renderDocument(this.core);
	}

	disconnectedCallback() {
		const { core } = this;
		core.collab.presence?.leave();
		core.detachedState = core.view?.state;
		core.view?.destroy();
		core.view = undefined;
		core.imageMedia.release();
	}

	async load(input: Uint8Array | ArrayBuffer): Promise<void> {
		const { core } = this;
		this.assertDocumentReplaceable();
		const generation = ++core.loadGeneration;
		try {
			const session = await loadDocument(input);
			if (generation !== core.loadGeneration) return;
			core.imageMedia.release();
			this.replaceDocument(session.model, session, false);
			core.shell.chrome?.setSaveState('saved');
			if (this.isConnected) renderDocument(core);
		} catch (cause) {
			const error = cause instanceof Error ? cause : new Error(String(cause));
			if (generation === core.loadGeneration) emit(this, 'document-error', error);
			throw error;
		}
	}

	setLoadedDocument(session: LoadedDocument) {
		this.assertDocumentReplaceable();
		this.replaceDocument(session.model, session);
		if (this.isConnected) renderDocument(this.core);
	}

	/** Join only after loading the authority's matching document snapshot and version. */
	startCollaboration(config: CollaborationConfig): void {
		const { core } = this;
		if (core.collab.active) throw new Error('Stop the current collaboration session first.');
		if (!core.view) throw new Error('Mount and load the document before starting collaboration.');
		core.collab.start(config);
		core.loadGeneration++;
		core.detachedState = undefined;
		renderDocument(core);
	}

	getPendingCollaboration(): StepBatch | null {
		return this.core.pendingCollaboration();
	}

	receiveCollaboration(batch: unknown): ClientReceiveResult['status'] {
		const { collab, view } = this.core;
		if (!collab.client || !view) throw new Error('No mounted collaboration session.');
		const result = collab.client.receive(view.state, batch);
		if (result.status === 'applied') this.core.applyTransaction(result.transaction, true);
		return result.status;
	}

	/** Stopping with pending edits requires an explicit discard of the transport queue. */
	stopCollaboration(discardPending = false): void {
		const { core } = this;
		const state: EditorState | undefined = core.view?.state ?? core.detachedState;
		if (!discardPending && state && core.collab.client?.pendingStepCount(state))
			throw new Error(
				'Acknowledge pending collaboration edits before stopping, or explicitly discard the queue.',
			);
		core.collab.stop();
		core.detachedState = undefined;
		if (this.isConnected) renderDocument(core);
	}

	private assertDocumentReplaceable() {
		if (this.core.collab.client)
			throw new Error('Stop collaboration before replacing the document.');
	}

	/** Swaps in a new model (and the session that produced it) and invalidates in-flight loads. */
	private replaceDocument(model: DocumentModel, session?: LoadedDocument, bumpGeneration = true) {
		const { core } = this;
		core.detachedState = undefined;
		if (bumpGeneration) core.loadGeneration++;
		core.loaded = session;
		core.model = model;
		core.inserts.reset();
		core.formatDialogs.closeAll();
		core.dirtyState.set(false);
	}
}

declare global {
	interface HTMLElementTagNameMap {
		'docx-editor': DocxEditorElement;
	}
}

export function registerDocxEditor() {
	if (typeof customElements === 'undefined' || typeof HTMLElement === 'undefined') return;
	if (!customElements.get('docx-editor')) customElements.define('docx-editor', DocxEditorElement);
}

registerDocxEditor();
