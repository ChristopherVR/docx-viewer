import { icon, iconButton } from './chrome-icons';
import { translateUiText } from './localization';

export interface StatusBarHandlers {
	setViewMode(mode: 'draft' | 'print'): void;
	setZoom(percent: number): void;
	showCompatibilityNotes(): void;
}

export interface StatusBar {
	element: HTMLElement;
	setPageAndWords(page: string, words: string): void;
	setLanguage(language: string): void;
	setNoteCount(count: number): void;
	setViewMode(mode: 'draft' | 'print'): void;
	setZoom(percent: number): void;
}

export const ZOOM_MIN = 50;
export const ZOOM_MAX = 200;

/** Word-style status bar: page/words/language on the left, views and zoom on the right. */
export function createStatusBar(handlers: StatusBarHandlers): StatusBar {
	const element = document.createElement('footer');
	element.className = 'dve-status';
	const left = document.createElement('div');
	left.className = 'dve-status-left';
	const page = document.createElement('span');
	page.className = 'dve-status-page';
	const words = document.createElement('span');
	words.className = 'dve-status-words';
	const language = document.createElement('span');
	language.className = 'dve-status-language';
	const notes = document.createElement('button');
	notes.type = 'button';
	notes.className = 'dve-status-notes';
	notes.hidden = true;
	const notesText = document.createElement('span');
	notes.append(icon('warning'), notesText);
	notes.addEventListener('click', () => handlers.showCompatibilityNotes());
	left.append(page, words, language, notes);

	const right = document.createElement('div');
	right.className = 'dve-status-right';
	const webView = iconButton('webView', 'Web Layout');
	const printView = iconButton('pageView', 'Print Layout');
	webView.addEventListener('click', () => handlers.setViewMode('draft'));
	printView.addEventListener('click', () => handlers.setViewMode('print'));
	const zoomOut = iconButton('minus', 'Zoom out');
	const zoomIn = iconButton('plus', 'Zoom in');
	const slider = document.createElement('input');
	slider.type = 'range';
	slider.className = 'dve-zoom-slider';
	slider.min = String(ZOOM_MIN);
	slider.max = String(ZOOM_MAX);
	slider.step = '10';
	slider.setAttribute('aria-label', 'Zoom level');
	const percent = document.createElement('button');
	percent.type = 'button';
	percent.className = 'dve-zoom-percent';
	percent.setAttribute('aria-label', 'Reset zoom');
	percent.title = 'Reset zoom';
	const clamp = (value: number) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, value));
	zoomOut.addEventListener('click', () => handlers.setZoom(clamp(Number(slider.value) - 10)));
	zoomIn.addEventListener('click', () => handlers.setZoom(clamp(Number(slider.value) + 10)));
	slider.addEventListener('input', () => handlers.setZoom(clamp(Number(slider.value))));
	percent.addEventListener('click', () => handlers.setZoom(100));
	right.append(webView, printView, zoomOut, slider, zoomIn, percent);
	element.append(left, right);

	return {
		element,
		setPageAndWords(pageText, wordText) {
			page.textContent = pageText;
			words.textContent = wordText;
		},
		setLanguage(value) {
			language.textContent = value;
			language.hidden = !value;
		},
		setNoteCount(count) {
			notes.hidden = count === 0;
			notesText.textContent = translateUiText(
				element,
				count === 1 ? 'status.note' : 'status.notes',
			).replace('{count}', String(count));
		},
		setViewMode(mode) {
			webView.setAttribute('aria-pressed', String(mode === 'draft'));
			printView.setAttribute('aria-pressed', String(mode === 'print'));
		},
		setZoom(value) {
			slider.value = String(value);
			percent.textContent = `${value}%`;
		},
	};
}
