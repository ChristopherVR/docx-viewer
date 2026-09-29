import type { EditorView } from 'prosemirror-view';
import { countWords } from './word-count';
import { localeOf, translate, type EditorLocale } from './localization';

export interface DocumentCounts {
	words: number;
	charactersWithSpaces: number;
	charactersWithoutSpaces: number;
	paragraphs: number;
}

/** Counts the selection when there is one, else the whole document, as Word's Word Count does. */
export function documentCounts(view: EditorView, language?: string): DocumentCounts {
	const { from, to, empty } = view.state.selection;
	const text = empty
		? view.state.doc.textBetween(0, view.state.doc.content.size, '\n', ' ')
		: view.state.doc.textBetween(from, to, '\n', ' ');
	let paragraphs = 0;
	view.state.doc.nodesBetween(
		empty ? 0 : from,
		empty ? view.state.doc.content.size : to,
		(node) => {
			if (node.type.name === 'paragraph' && node.textContent.trim()) paragraphs++;
		},
	);
	return {
		words: countWords(text, language),
		charactersWithSpaces: [...text.replace(/\n/g, '')].length,
		charactersWithoutSpaces: [...text.replace(/\s/g, '')].length,
		paragraphs,
	};
}

let close: (() => void) | undefined;

/** Shows the counts in a small dismissible panel under `anchor`. */
export function showWordCount(
	anchor: HTMLElement,
	view: EditorView,
	locale: EditorLocale = localeOf(anchor),
	language?: string,
): void {
	close?.();
	const counts = documentCounts(view, language);
	const panel = document.createElement('div');
	panel.className = 'ribbon-popover dve-word-count';
	panel.setAttribute('role', 'dialog');
	panel.setAttribute('aria-label', translate(locale, 'Word count'));
	const number = new Intl.NumberFormat(locale);
	const rows: Array<[Parameters<typeof translate>[1], number]> = [
		['Words', counts.words],
		['Characters (no spaces)', counts.charactersWithoutSpaces],
		['Characters (with spaces)', counts.charactersWithSpaces],
		['Paragraphs', counts.paragraphs],
	];
	for (const [label, value] of rows) {
		const name = document.createElement('span');
		name.textContent = translate(locale, label);
		const amount = document.createElement('strong');
		amount.textContent = number.format(value);
		panel.append(name, amount);
	}
	const box = anchor.getBoundingClientRect();
	panel.style.left = `${Math.max(4, box.left)}px`;
	panel.style.top = `${box.bottom + 2}px`;
	const root = anchor.getRootNode();
	(root instanceof ShadowRoot ? root : document.body).append(panel);
	const outside = (event: Event) => {
		if (!event.composedPath().includes(panel)) close?.();
	};
	const escape = (event: KeyboardEvent) => {
		if (event.key === 'Escape') close?.();
	};
	document.addEventListener('pointerdown', outside, true);
	document.addEventListener('keydown', escape, true);
	close = () => {
		document.removeEventListener('pointerdown', outside, true);
		document.removeEventListener('keydown', escape, true);
		panel.remove();
		close = undefined;
	};
}
