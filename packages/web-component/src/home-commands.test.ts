// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { changeCase, transformCase } from './change-case';
import { isFormatPainterArmed, toggleFormatPainter } from './format-painter';
import { createRibbon } from './ribbon';
import { schema } from './schema';

function editor(text: string, marks = false) {
	const content = marks
		? [schema.text('bold', [schema.marks.bold!.create()]), schema.text(text)]
		: [schema.text(text)];
	const doc = schema.node('doc', null, [schema.nodes.paragraph!.create({ id: 'p' }, content)]);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const select = (view: EditorView, from: number, to: number) =>
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));
const textOf = (view: EditorView) => view.state.doc.textContent;

describe('transformCase', () => {
	it('handles every Word mode', () => {
		expect(transformCase('hello WORLD. this is it! ok', 'sentence')).toBe(
			'Hello world. This is it! Ok',
		);
		expect(transformCase('hello world', 'lower')).toBe('hello world');
		expect(transformCase('Hello world', 'upper')).toBe('HELLO WORLD');
		expect(transformCase("it's a dog-eat-dog world", 'title')).toBe("It's A Dog-Eat-Dog World");
		expect(transformCase('Hello World', 'toggle')).toBe('hELLO wORLD');
	});

	it('leaves a mid-sentence fragment lowercase when it does not start a sentence', () => {
		expect(transformCase('and more', 'sentence', false)).toBe('and more');
	});
});

describe('changeCase', () => {
	it('changes only the selection and keeps run marks', () => {
		const view = editor(' and more', true);
		select(view, 1, 5);
		expect(changeCase(view, 'upper')).toBe(true);
		expect(textOf(view)).toBe('BOLD and more');
		expect(view.state.doc.firstChild!.firstChild!.marks[0]!.type.name).toBe('bold');
	});

	it('works on the word at a collapsed caret', () => {
		const view = editor('alpha beta gamma');
		select(view, 9, 9);
		changeCase(view, 'upper');
		expect(textOf(view)).toBe('alpha BETA gamma');
	});

	it('reports no change when nothing differs or there is no word', () => {
		const view = editor('ABC');
		select(view, 1, 4);
		expect(changeCase(view, 'upper')).toBe(false);
	});
});

describe('format painter', () => {
	it('copies character marks to the next selection, then disarms', async () => {
		const view = editor(' plain', true);
		const states: boolean[] = [];
		select(view, 1, 3);
		expect(toggleFormatPainter(view, (on) => states.push(on))).toBe(true);
		expect(isFormatPainterArmed(view)).toBe(true);
		select(view, 5, 8);
		view.dom.dispatchEvent(new MouseEvent('mouseup'));
		await new Promise((resolve) => setTimeout(resolve, 0));
		const painted = view.state.doc.nodeAt(6)!;
		expect(painted.marks.map((m) => m.type.name)).toContain('bold');
		expect(isFormatPainterArmed(view)).toBe(false);
		expect(states).toEqual([true, false]);
	});

	it('is cancelled by Escape and by a second click', () => {
		const view = editor('abc');
		select(view, 1, 2);
		toggleFormatPainter(view, () => {});
		view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(isFormatPainterArmed(view)).toBe(false);
		toggleFormatPainter(view, () => {});
		toggleFormatPainter(view, () => {});
		expect(isFormatPainterArmed(view)).toBe(false);
	});
});

describe('Home ribbon commands', () => {
	it('emits changeCase, formatPainter, showMarks, selectAll and replace actions', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		for (const name of ['Format painter', 'Show paragraph marks', 'Select all', 'Replace'])
			ribbon.querySelector<HTMLButtonElement>(`button[aria-label="${name}"]`)!.click();
		const menu = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Change case"]')!;
		menu.value = 'upper';
		menu.dispatchEvent(new Event('change'));
		expect(seen).toEqual([
			{ type: 'formatPainter' },
			{ type: 'showMarks' },
			{ type: 'selectAll' },
			{ type: 'search', focus: 'replace' },
			{ type: 'changeCase', value: 'upper' },
		]);
		expect(menu.selectedIndex).toBe(-1);
	});
});
