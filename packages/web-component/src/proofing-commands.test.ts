// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { fitZoomPercent } from './zoom-fit';
import { formatDateTime, insertPlainText, SYMBOLS } from './insert-text-commands';
import { documentCounts, showWordCount } from './word-count-panel';
import { createRibbon } from './ribbon';
import { schema } from './schema';
import { syncSpellingButton } from './spelling';

const editor = (...paragraphs: string[]) => {
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((text, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}` }, text ? schema.text(text) : undefined),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
};

afterEach(() => (document.body.innerHTML = ''));

describe('fitZoomPercent', () => {
	const page = { width: 816, height: 1056 };
	it('returns 100 for actual size and guards bad page sizes', () => {
		expect(fitZoomPercent('actual', { width: 400, height: 300 }, page)).toBe(100);
		expect(fitZoomPercent('width', { width: 400, height: 300 }, { width: 0, height: 0 })).toBe(100);
	});
	it("fits the width, or the whole page, and clamps to Word's range", () => {
		expect(fitZoomPercent('width', { width: 1632, height: 300 }, page)).toBe(200);
		expect(fitZoomPercent('page', { width: 1632, height: 528 }, page)).toBe(50);
		expect(fitZoomPercent('width', { width: 10, height: 10 }, page)).toBe(10);
		expect(fitZoomPercent('width', { width: 99999, height: 10 }, page)).toBe(500);
	});
});

describe('insert text commands', () => {
	it('inserts a symbol at the caret and replaces a selection', () => {
		const view = editor('ab');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2, 2)));
		expect(insertPlainText(view, SYMBOLS[0])).toBe(true);
		expect(view.state.doc.textContent).toBe('a©b');
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 3)));
		insertPlainText(view, '—');
		expect(view.state.doc.textContent).toBe('—b');
	});

	it('formats the date for the display locale', () => {
		const now = new Date(Date.UTC(2026, 8, 29, 12, 0));
		expect(formatDateTime('long', 'en-US', now)).toBe('September 29, 2026');
		expect(formatDateTime('long', 'de', now)).toBe('29. September 2026');
		expect(formatDateTime('long', 'not a locale!', now)).toBe('September 29, 2026');
	});

	it('does not edit a read-only view', () => {
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc: editor('x').state.doc, schema }),
			editable: () => false,
		});
		expect(insertPlainText(view, '©')).toBe(false);
	});
});

describe('word count', () => {
	it('counts the document, or only the selection when there is one', () => {
		const view = editor('one two three', '', 'four five');
		expect(documentCounts(view)).toEqual({
			words: 5,
			charactersWithSpaces: 22,
			charactersWithoutSpaces: 19,
			paragraphs: 2,
		});
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4)));
		expect(documentCounts(view).words).toBe(1);
	});

	it('shows the counts in a panel that Escape dismisses', () => {
		const view = editor('alpha beta');
		const button = document.createElement('button');
		document.body.append(button);
		showWordCount(button, view, 'fr');
		const panel = document.querySelector('.dve-word-count')!;
		expect(panel.getAttribute('aria-label')).toBe('Nombre de mots');
		expect(panel.textContent).toContain('Mots');
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(document.querySelector('.dve-word-count')).toBeNull();
	});
});

describe('spelling and the new ribbon controls', () => {
	it('reflects the surface spellcheck state on the Spelling button', () => {
		const ribbon = createRibbon();
		const view = editor('x');
		view.dom.spellcheck = false;
		syncSpellingButton(ribbon, view);
		const button = ribbon.querySelector('[aria-label="Spelling"]')!;
		expect(button.getAttribute('aria-pressed')).toBe('false');
		view.dom.spellcheck = true;
		syncSpellingButton(ribbon, view);
		expect(button.getAttribute('aria-pressed')).toBe('true');
	});

	it('emits symbol, date, zoom-fit, word-count and spelling actions', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const symbol = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Symbol"]')!;
		symbol.value = '€';
		symbol.dispatchEvent(new Event('change'));
		const date = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Date and time"]')!;
		date.value = 'time';
		date.dispatchEvent(new Event('change'));
		for (const name of ['Page width', 'One page', 'Word count', 'Spelling'])
			ribbon.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!.click();
		expect(seen).toEqual([
			{ type: 'insertSymbol', value: '€' },
			{ type: 'insertDateTime', value: 'time' },
			{ type: 'zoomFit', mode: 'width' },
			{ type: 'zoomFit', mode: 'page' },
			{ type: 'wordCount' },
			{ type: 'spelling' },
		]);
	});
});
