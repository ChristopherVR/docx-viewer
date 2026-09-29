// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { runRibbonCommand } from './editor-commands';
import { headingStyleId, setHeadingLevel } from './heading-commands';
import { readAloudAvailable, textToRead, toggleReadAloud } from './read-aloud';
import { createRibbon } from './ribbon';
import { schema } from './schema';

const model = createDocument();

function editor(...texts: string[]) {
	const doc = schema.node(
		'doc',
		null,
		texts.map((text, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}` }, schema.text(text)),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const select = (view: EditorView, from: number, to: number) =>
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));

afterEach(() => {
	vi.unstubAllGlobals();
});

describe('References > Add text', () => {
	it('finds the heading styles of a new document by name or id', () => {
		expect(headingStyleId(model, 1)).toBeDefined();
		expect(headingStyleId(model, 3)).toBeDefined();
	});

	it('applies a heading level to the selected paragraphs and clears it with level 0', () => {
		const view = editor('One', 'Two');
		select(view, 1, view.state.doc.content.size - 1);
		expect(setHeadingLevel(view, model, 2)).toBe(true);
		const style = headingStyleId(model, 2);
		view.state.doc.forEach((p) => expect(p.attrs.style).toBe(style));
		expect(setHeadingLevel(view, model, 0)).toBe(true);
		view.state.doc.forEach((p) => expect(p.attrs.style).toBe(''));
	});

	it('reports failure when the document has no such heading style', () => {
		const view = editor('One');
		const bare = createDocument();
		bare.paragraphStyles = { docDefaults: {}, warnings: [], styles: {} };
		expect(setHeadingLevel(view, bare, 1)).toBe(false);
	});
});

describe('Blank page', () => {
	it('inserts two page breaks', () => {
		const view = editor('Start');
		select(view, 3, 3);
		runRibbonCommand(view, { type: 'blankPage' });
		let breaks = 0;
		view.state.doc.descendants((node) => {
			if (node.type.name === 'pageBreak' || node.attrs.pageBreakBefore) breaks++;
			if (node.type.name.toLowerCase().includes('break') && node.type.name !== 'hardBreak')
				breaks++;
		});
		expect(breaks).toBeGreaterThanOrEqual(2);
	});
});

describe('Read aloud', () => {
	it('reads the selection, or from the caret to the end', () => {
		const view = editor('alpha beta', 'gamma');
		select(view, 1, 6);
		expect(textToRead(view)).toBe('alpha');
		select(view, 7, 7);
		expect(textToRead(view)).toBe('beta\ngamma');
	});

	it('is unavailable, not faked, without speech synthesis', () => {
		vi.stubGlobal('speechSynthesis', undefined);
		expect(readAloudAvailable()).toBe(false);
		expect(toggleReadAloud('text', () => {})).toBe(false);
	});

	it('speaks, reports state, and stops on the second call', () => {
		const spoken: string[] = [];
		const listeners: Record<string, () => void> = {};
		vi.stubGlobal('speechSynthesis', {
			speak: (u: { text: string }) => spoken.push(u.text),
			cancel: vi.fn(),
		});
		vi.stubGlobal(
			'SpeechSynthesisUtterance',
			class {
				constructor(public text: string) {}
				addEventListener(name: string, fn: () => void) {
					listeners[name] = fn;
				}
			},
		);
		const states: boolean[] = [];
		expect(toggleReadAloud('hello there', (on) => states.push(on))).toBe(true);
		expect(spoken).toEqual(['hello there']);
		expect(toggleReadAloud('ignored', (on) => states.push(on))).toBe(false);
		expect(states).toEqual([true, false]);
		toggleReadAloud('again', (on) => states.push(on));
		listeners.end?.();
		expect(states).toEqual([true, false, true, false]);
	});

	it('does not start for empty text', () => {
		vi.stubGlobal('speechSynthesis', { speak: vi.fn(), cancel: vi.fn() });
		vi.stubGlobal('SpeechSynthesisUtterance', class {});
		expect(toggleReadAloud('', () => {})).toBe(false);
	});
});

describe('new ribbon controls', () => {
	it('emit their actions', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		for (const name of ['Blank page', 'Read aloud', 'Gridlines'])
			ribbon.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!.click();
		const menu = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Add text"]')!;
		menu.value = '2';
		menu.dispatchEvent(new Event('change'));
		expect(seen).toEqual([
			{ type: 'blankPage' },
			{ type: 'readAloud' },
			{ type: 'gridlines' },
			{ type: 'addText', level: 2 },
		]);
	});
});
