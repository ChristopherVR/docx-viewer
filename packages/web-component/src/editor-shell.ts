import { applyShellLabels } from './shell-labels';
import type { DocumentModel } from '@christophervr/docx-core';
import { createDocument } from '@christophervr/docx-core';
import type { EditorCore } from './editor-core';
import { emit, on } from './events';
import { createRibbon } from './ribbon';
import { editorStyleText } from './styles';
import { runRibbonCommand } from './editor-commands';
import { createSearchPanel } from './search-panel';
import { createPrintLayoutController } from './print-layout-view';
import { moveCursorToBlock } from './print-layout-cursor';
import { ReviewController } from './review-controller';
import { EditorChrome } from './editor-chrome';
import { dispatchDocumentError } from './editor-host';
import { routeRibbonAction } from './ribbon-router';
import { countWords } from './word-count';
import { PageNavigator } from './page-navigator';
import { applyViewOptions } from './view-options';
import { attachEditorInteractions } from './editor-interactions';

/** Element-level operations the chrome needs; they stay on the element because they are public API. */
export interface ShellApi {
	setReadOnly(readOnly: boolean): void;
	setDocumentModel(model: DocumentModel): void;
	load(bytes: Uint8Array | ArrayBuffer): Promise<void>;
	saveBytes(): Promise<Uint8Array>;
}

function createChrome(core: EditorCore, api: ShellApi): EditorChrome {
	const { element, shell } = core;
	return new EditorChrome({
		element,
		model: () => core.model,
		ribbon: () => shell.toolbar,
		wordCount: () => {
			const doc = core.view?.state.doc;
			return doc
				? countWords(doc.textBetween(0, doc.content.size, ' ').trim(), element.lang || undefined)
				: 0;
		},
		locale: () => core.locale,
		readOnly: () => core.readOnly,
		setReadOnly: (readOnly) => {
			api.setReadOnly(readOnly);
			emit(element, 'readonly-change', readOnly);
		},
		newDocument: () => api.setDocumentModel(createDocument()),
		load: (bytes) => api.load(bytes),
		save: () => api.saveBytes(),
		saveStateChanged: (state) => {
			if (state !== 'saving') core.dirtyState.set(state === 'dirty');
		},
		print: () => core.pages.print(),
		history: (key) => {
			if (!core.view) return;
			runRibbonCommand(core.view, { type: 'history', key }, core.collab.ids);
			if (typeof document.execCommand === 'function') core.view.focus();
		},
		toggleComments: () => shell.review?.handleComments('toggle'),
		setViewMode: (mode) => core.pages.setViewMode(mode),
		setZoom: (percent) => core.pages.setZoom(percent),
		reportError: (error) => dispatchDocumentError(element, error),
	});
}

/** Builds the shadow-DOM shell (ribbon, canvas, panels, chrome) once; later calls are no-ops. */
export function buildShell(core: EditorCore, api: ShellApi): void {
	const { element, shell } = core;
	if (shell.toolbar) return;
	element.classList.add('dve-host');
	const root = element.attachShadow({ mode: 'open' });
	const style = document.createElement('style');
	style.textContent = editorStyleText;
	const frame = document.createElement('section');
	frame.className = 'dve-frame';
	const toolbar = createRibbon(core.locale);
	on(toolbar, 'ribbon-action', (event) => routeRibbonAction(core, event.detail));
	const canvas = document.createElement('main');
	canvas.className = 'dve-canvas';
	const paper = document.createElement('div');
	paper.className = 'dve-paper';
	canvas.append(paper);
	const printLayout = createPrintLayoutController(
		canvas,
		(blockId, offset) => {
			core.pages.setViewMode('draft');
			if (core.view) moveCursorToBlock(core.view, blockId, offset);
		},
		(partName, contentType) => core.imageMedia.urlFor(partName, contentType),
		() => core.refreshControls(),
	);
	canvas.append(printLayout.element);
	canvas.addEventListener('scroll', () => {
		if (core.pages.viewMode === 'print') {
			printLayout.refreshCurrentPage();
			core.refreshControls();
		}
	});
	shell.printLayout = printLayout;
	shell.navigator = new PageNavigator({
		pages: () => printLayout.pageElements(),
		version: () => printLayout.layoutVersion(),
		goTo: (page) => {
			printLayout.scrollToPage(page);
			core.refreshControls();
		},
		close: () => element.toggleAttribute('show-thumbnails', false),
	});
	shell.searchPanel = createSearchPanel({
		getView: () => core.view,
		onClose: () => core.view?.focus(),
	});
	shell.searchPanel.setLocale(core.locale);
	const review = new ReviewController({
		getModel: () => core.model,
		setModel: (model) => {
			core.model = model;
		},
		getView: () => core.view,
		getReviewAuthor: () => core.reviewAuthor,
		getCollaborationIds: () => core.collab.ids,
		notifyChange: () => core.notifyChange(),
		refresh: () => core.refreshControls(),
	});
	shell.review = review;
	review.setLocale(core.locale);
	core.inserts.setLocale(core.locale);
	const body = document.createElement('div');
	body.className = 'dve-body';
	body.append(shell.navigator.element, canvas, review.commentsPanel.element);
	frame.append(toolbar, shell.searchPanel.element, body);
	const chrome = createChrome(core, api);
	shell.chrome = chrome;
	chrome.mount(frame, toolbar);
	frame.append(
		core.inserts.linkDialog.element,
		core.inserts.pictureDialog.element,
		core.inserts.pictureInput,
	);
	if (core.pendingFileName) chrome.fileName = core.pendingFileName;
	root.append(style, frame);
	Object.assign(shell, { toolbar, canvas, paper });
	// After the toolbar is registered: the File tab label is localized through the shell.
	chrome.setLocale(core.locale);
	applyViewOptions(core);
	attachEditorInteractions(core, frame);
	element.setAttribute('role', 'region');
	applyShellLabels(element, paper, core.locale);
}
