import { saveDocx, type DocumentModel } from '@christophervr/docx-core';
import { createBackstage, type Backstage } from './backstage';
import {
	announceFileCommand,
	downloadBytes,
	downloadText,
	withExtension,
	type FileCommand,
} from './file-commands';
import { documentStats, plainText } from './document-stats';
import { localizeElement, translate, type EditorLocale } from './localization';
import { createStatusBar, type StatusBar } from './status-bar';
import { createTitleBar, type SaveState, type TitleBar } from './title-bar';

/** What the chrome needs from the editor element; keeps component.ts free of chrome details. */
export interface ChromeHost {
	element: HTMLElement;
	model(): DocumentModel;
	ribbon(): HTMLElement | undefined;
	wordCount(): number;
	locale(): EditorLocale;
	readOnly(): boolean;
	setReadOnly(readOnly: boolean): void;
	newDocument(): void;
	load(bytes: Uint8Array): Promise<void>;
	save(): Promise<Uint8Array>;
	print(): void;
	history(key: 'undo' | 'redo'): void;
	toggleComments(): void;
	setViewMode(mode: 'draft' | 'print'): void;
	setZoom(percent: number): void;
	reportError(error: Error): void;
	/** Mirrors the title bar's save state into `element.dirty`. */
	saveStateChanged?(state: SaveState): void;
	/** Editor options shown on File > Options. */
	options(): { locale: string; theme: string; author: string };
	setOption(key: 'locale' | 'theme' | 'author', value: string): void;
}

export const DEFAULT_FILE_NAME = 'Document1.docx';

/** Title bar, File tab/backstage, status bar and the default browser file workflow. */
export class EditorChrome {
	readonly titleBar: TitleBar;
	readonly backstage: Backstage;
	readonly statusBar: StatusBar;
	readonly fileInput: HTMLInputElement;
	private _fileName = DEFAULT_FILE_NAME;
	private state: SaveState = 'saved';

	constructor(private readonly host: ChromeHost) {
		this.titleBar = createTitleBar({
			fileCommand: (command) => void this.run(command),
			history: (key) => host.history(key),
			toggleComments: () => host.toggleComments(),
			setReadOnly: (readOnly) => host.setReadOnly(readOnly),
			ribbon: () => host.ribbon(),
		});
		this.backstage = createBackstage({
			fileCommand: (command, fileName) => void this.run(command, fileName),
			close: () => this.closeBackstage(),
			summary: () => ({
				fileName: this._fileName,
				model: host.model(),
				words: host.wordCount(),
				stats: documentStats(host.model()),
				saveState: this.state,
			}),
			options: () => host.options(),
			setOption: (key, value) => host.setOption(key, value),
		});
		this.statusBar = createStatusBar({
			setViewMode: (mode) => host.setViewMode(mode),
			setZoom: (percent) => host.setZoom(percent),
			showCompatibilityNotes: () => this.backstage.open('info'),
		});
		this.fileInput = document.createElement('input');
		this.fileInput.type = 'file';
		this.fileInput.accept = '.docx,.doc';
		this.fileInput.hidden = true;
		this.fileInput.className = 'dve-file-input';
		this.fileInput.addEventListener('change', () => void this.openSelectedFile());
		this.titleBar.setFileName(this._fileName);
		this.titleBar.setSaveState(this.state);
		this.statusBar.setViewMode('draft');
		this.statusBar.setZoom(100);
	}

	/** Adds the chrome around an existing frame: title bar first, status bar and backstage last. */
	mount(frame: HTMLElement, ribbon: HTMLElement): void {
		frame.prepend(this.titleBar.element);
		frame.append(this.statusBar.element, this.backstage.element, this.fileInput);
		const tabs = ribbon.querySelector('.ribbon-tabs');
		const fileTab = document.createElement('button');
		fileTab.type = 'button';
		fileTab.className = 'dve-file-tab';
		fileTab.textContent = 'File';
		fileTab.setAttribute('aria-haspopup', 'dialog');
		fileTab.addEventListener('click', () => this.backstage.open('info'));
		tabs?.prepend(fileTab);
	}

	get fileName(): string {
		return this._fileName;
	}
	set fileName(value: string) {
		this._fileName = value || DEFAULT_FILE_NAME;
		this.titleBar.setFileName(this._fileName);
	}
	get isDirty(): boolean {
		return this.state === 'dirty';
	}

	setSaveState(state: SaveState): void {
		this.state = state;
		this.host.saveStateChanged?.(state);
		this.titleBar.setSaveState(state);
	}

	setLocale(locale: EditorLocale): void {
		for (const element of [this.titleBar.element, this.backstage.element, this.statusBar.element]) {
			element.dataset.editorLocale = locale;
			localizeElement(element, locale);
		}
		const fileTab = this.host.ribbon()?.querySelector('.dve-file-tab');
		if (fileTab) fileTab.textContent = translate(locale, 'File');
		this.fileInput.setAttribute('aria-label', translate(locale, 'Open'));
		this.titleBar.setSaveState(this.state);
	}

	refresh(pageText: string, wordText: string): void {
		const model = this.host.model();
		this.statusBar.setPageAndWords(pageText, wordText);
		this.statusBar.setNoteCount(model.warnings.length);
		this.statusBar.setLanguage(this.host.element.lang || '');
		this.titleBar.setReadOnly(this.host.readOnly());
	}

	closeBackstage(): void {
		this.backstage.close();
		this.host.ribbon()?.querySelector<HTMLButtonElement>('.dve-file-tab')?.focus();
	}

	/** Runs a file command unless a host cancels the `file-command` event to handle it itself. */
	async run(command: FileCommand, fileName?: string): Promise<void> {
		if (!announceFileCommand(this.host.element, command, fileName)) return;
		try {
			if (command === 'new') this.newDocument();
			else if (command === 'open') this.fileInput.click();
			else if (command === 'print') this.host.print();
			else if (command === 'exportText')
				downloadText(plainText(this.host.model()), withExtension(this._fileName, 'txt'));
			else if (command === 'saveAs') {
				const name = (fileName ?? '').trim();
				if (!name) return;
				// The bytes are in the format the document was opened in, so the extension follows it.
				const extension = this._fileName.split('.').pop()?.toLowerCase() === 'doc' ? 'doc' : 'docx';
				this.fileName = withExtension(name, extension);
				await this.run('save');
			} else if (command === 'save') {
				this.setSaveState('saving');
				downloadBytes(await this.host.save(), this._fileName);
				this.setSaveState('saved-local');
			} else if (command === 'export')
				downloadBytes(
					await saveDocx(structuredClone(this.host.model())),
					withExtension(this._fileName, 'docx'),
				);
		} catch (cause) {
			if (command === 'save') this.setSaveState('dirty');
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
		}
	}

	private newDocument(): void {
		const prompt = translate(
			this.host.locale(),
			'Discard unsaved changes and start a new document?',
		);
		if (this.isDirty && typeof confirm === 'function' && !confirm(prompt)) return;
		this.host.newDocument();
		this.fileName = DEFAULT_FILE_NAME;
		this.setSaveState('saved');
	}

	private async openSelectedFile(): Promise<void> {
		const file = this.fileInput.files?.[0];
		if (!file) return;
		try {
			await this.host.load(new Uint8Array(await file.arrayBuffer()));
			this.fileName = file.name;
			this.setSaveState('saved');
		} catch (cause) {
			this.host.reportError(cause instanceof Error ? cause : new Error(String(cause)));
		} finally {
			this.fileInput.value = '';
		}
	}
}
