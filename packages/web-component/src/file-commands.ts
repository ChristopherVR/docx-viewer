/**
 * File commands raised by the editor chrome (title bar and File backstage). Each command is first
 * announced as a cancelable `file-command` event so hosts can take over file I/O; when no host
 * cancels it, the editor performs a browser-only default (file picker, download, print).
 */
import { emit } from './events';

export type FileCommand = 'new' | 'open' | 'save' | 'export' | 'print';

export interface FileCommandDetail {
	command: FileCommand;
}

export function announceFileCommand(host: HTMLElement, command: FileCommand): boolean {
	return emit(host, 'file-command', { command }, { cancelable: true });
}

const WORD_TYPES: Record<string, string> = {
	doc: 'application/msword',
	docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

/** Wraps saved bytes in a Blob typed from the file name's extension (docx unless `.doc`). */
export function wordBlob(bytes: Uint8Array, fileName: string): Blob {
	const extension = fileName.split('.').pop()?.toLowerCase() ?? 'docx';
	return new Blob([new Uint8Array(bytes)], { type: WORD_TYPES[extension] ?? WORD_TYPES.docx });
}

/** Starts a browser download of `bytes` named `fileName`. */
export function downloadBytes(bytes: Uint8Array, fileName: string): void {
	const url = URL.createObjectURL(wordBlob(bytes, fileName));
	const anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = fileName;
	anchor.style.display = 'none';
	document.body.append(anchor);
	anchor.click();
	anchor.remove();
	setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** Replaces the extension of `fileName` (or appends one). */
export function withExtension(fileName: string, extension: string): string {
	const base = fileName.replace(/\.[^./\\]+$/, '') || 'Document';
	return `${base}.${extension}`;
}
