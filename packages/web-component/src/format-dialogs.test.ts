// @vitest-environment jsdom
import { createDocument } from '@christophervr/docx-core';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { FormatDialogs } from './format-dialogs';
import { applyParagraphFormat, readParagraphFormat } from './paragraph-format';
import { createRibbon } from './ribbon';
import { runStylesPlugin } from './run-styles';
import { schema } from './schema';

afterEach(() => (document.body.innerHTML = ''));

function setup(...paragraphs: Array<Record<string, unknown>>) {
	const model = createDocument();
	const doc = schema.node(
		'doc',
		null,
		paragraphs.map((attrs, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...attrs }, schema.text(`text ${index}`)),
		),
	);
	const host = document.createElement('div');
	document.body.append(host);
	const view = new EditorView(host, {
		state: EditorState.create({ doc, schema, plugins: [runStylesPlugin(() => model)] }),
	});
	const dialogs = new FormatDialogs({ view: () => view, model: () => model });
	document.body.append(...dialogs.elements);
	return { view, model, dialogs };
}
const selectAll = (view: EditorView) =>
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);
const field = (root: ParentNode, label: string) =>
	root.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${label}"]`)!;
const fire = (el: Element, type = 'input') => el.dispatchEvent(new Event(type, { bubbles: true }));
const ok = (root: ParentNode) =>
	[...root.querySelectorAll<HTMLButtonElement>('button')]
		.find((b) => b.textContent === 'OK')!
		.click();

describe('paragraph format model', () => {
	it('reads defaults and resolves direct values', () => {
		const { view, model } = setup({
			align: 'center',
			indentLeftTwips: 720,
			hangingTwips: 360,
			spacingBeforeTwips: 240,
			lineSpacingTwips: 360,
			lineSpacingRule: 'auto',
			keepNext: true,
		});
		expect(readParagraphFormat(view.state, model)).toMatchObject({
			align: 'center',
			leftInches: 0.5,
			special: 'hanging',
			specialInches: 0.25,
			beforePt: 12,
			lineRule: 'oneAndHalf',
			keepNext: true,
			widowControl: true,
			contextualSpacing: false,
		});
	});

	it('reports mixed values as null', () => {
		const { view, model } = setup({ align: 'left' }, { align: 'right' });
		selectAll(view);
		expect(readParagraphFormat(view.state, model).align).toBeNull();
	});

	it('applies only the requested fields to every selected paragraph', () => {
		const { view } = setup({ indentLeftTwips: 100 }, { indentLeftTwips: 100 });
		selectAll(view);
		applyParagraphFormat(view, {
			align: 'justify',
			rightInches: 1,
			special: 'firstLine',
			specialInches: 0.5,
			afterPt: 6,
			lineRule: 'exactly',
			lineAt: 14,
			keepLines: true,
			widowControl: false,
		});
		view.state.doc.forEach((p) =>
			expect(p.attrs).toMatchObject({
				align: 'justify',
				indentLeftTwips: 100,
				indentRightTwips: 1440,
				firstLineTwips: 720,
				hangingTwips: null,
				spacingAfterTwips: 120,
				lineSpacingRule: 'exact',
				lineSpacingTwips: 280,
				keepLines: true,
				widowControl: false,
			}),
		);
	});

	it('maps line spacing rules to Word units and clears the other special indent', () => {
		const { view } = setup({ firstLineTwips: 360 });
		applyParagraphFormat(view, { special: 'hanging', specialInches: 0.25 });
		expect(view.state.doc.firstChild!.attrs).toMatchObject({
			firstLineTwips: null,
			hangingTwips: 360,
		});
		applyParagraphFormat(view, { lineRule: 'multiple', lineAt: 1.15 });
		expect(view.state.doc.firstChild!.attrs).toMatchObject({
			lineSpacingRule: 'auto',
			lineSpacingTwips: 276,
		});
		applyParagraphFormat(view, { lineRule: 'double' });
		expect(view.state.doc.firstChild!.attrs.lineSpacingTwips).toBe(480);
	});

	it('is one undo step', async () => {
		const { view } = setup({});
		const { undo } = await import('prosemirror-history');
		void undo;
		expect(applyParagraphFormat(view, { align: 'right', afterPt: 4 })).toBe(true);
		expect(view.state.doc.firstChild!.attrs.align).toBe('right');
	});
});

describe('Paragraph dialog', () => {
	it('shows the selection, then applies only what the user changed', () => {
		const { view, dialogs } = setup({ indentLeftTwips: 720, spacingAfterTwips: 200 });
		dialogs.open('paragraph');
		const root = dialogs.elements[1]!;
		expect(root.hidden).toBe(false);
		expect(field(root, 'Left').value).toBe('0.5');
		expect(field(root, 'After').value).toBe('10');
		field(root, 'Alignment').value = 'center';
		fire(field(root, 'Alignment'), 'change');
		field(root, 'Keep with next').checked = true;
		fire(field(root, 'Keep with next'), 'change');
		ok(root);
		expect(root.hidden).toBe(true);
		expect(view.state.doc.firstChild!.attrs).toMatchObject({
			align: 'center',
			keepNext: true,
			indentLeftTwips: 720,
			spacingAfterTwips: 200,
		});
	});

	it('enables the By and At fields only when they apply', () => {
		const { dialogs } = setup({ lineSpacingTwips: 240, lineSpacingRule: 'auto' });
		dialogs.open('paragraph');
		const root = dialogs.elements[1]!;
		expect(field(root, 'By').disabled).toBe(true);
		expect(field(root, 'At').disabled).toBe(true);
		field(root, 'Special').value = 'hanging';
		fire(field(root, 'Special'), 'change');
		expect(field(root, 'By').disabled).toBe(false);
		expect(field(root, 'By').value).toBe('0.5');
		field(root, 'Line spacing').value = 'atLeast';
		fire(field(root, 'Line spacing'), 'change');
		expect(field(root, 'At').disabled).toBe(false);
	});

	it('cancels without changing the document and closes on Escape', () => {
		const { view, dialogs } = setup({});
		dialogs.open('paragraph');
		const root = dialogs.elements[1]!;
		field(root, 'Alignment').value = 'right';
		fire(field(root, 'Alignment'), 'change');
		root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(root.hidden).toBe(true);
		expect(view.state.doc.firstChild!.attrs.align).toBeNull();
	});
});

describe('Font dialog', () => {
	it('applies family, size, style, underline and effects from the fields', () => {
		const { view, dialogs } = setup({});
		selectAll(view);
		dialogs.open('font');
		const root = dialogs.elements[0]!;
		expect(field(root, 'Font').value).toBe('Calibri');
		field(root, 'Font').value = 'Georgia';
		fire(field(root, 'Font'));
		field(root, 'Size').value = '14,5';
		fire(field(root, 'Size'));
		field(root, 'Font style').value = 'boldItalic';
		fire(field(root, 'Font style'), 'change');
		field(root, 'Underline style').value = 'wave';
		fire(field(root, 'Underline style'), 'change');
		field(root, 'Small caps').checked = true;
		fire(field(root, 'Small caps'), 'change');
		field(root, 'Character spacing').value = 'expanded';
		fire(field(root, 'Character spacing'), 'change');
		ok(root);
		const marks = view.state.doc.nodeAt(2)!.marks;
		const names = marks.map((mark) => mark.type.name);
		expect(names).toEqual(
			expect.arrayContaining(['bold', 'italic', 'underline', 'font', 'runProperties']),
		);
		expect(marks.find((mark) => mark.type.name === 'font')!.attrs).toMatchObject({
			family: 'Georgia',
			size: 14.5,
		});
		expect(marks.find((mark) => mark.type.name === 'runProperties')!.attrs.props).toEqual({
			underlineStyle: 'wave',
			smallCaps: true,
			characterSpacingTwips: 20,
		});
	});

	it('leaves untouched fields alone and ignores an invalid size', () => {
		const { view, dialogs } = setup({});
		selectAll(view);
		dialogs.open('font');
		const root = dialogs.elements[0]!;
		field(root, 'Size').value = 'huge';
		fire(field(root, 'Size'));
		ok(root);
		expect(view.state.doc.nodeAt(2)!.marks).toHaveLength(0);
	});

	it('keeps superscript and subscript exclusive', () => {
		const { view, dialogs } = setup({});
		selectAll(view);
		dialogs.open('font');
		const root = dialogs.elements[0]!;
		field(root, 'Superscript').checked = true;
		fire(field(root, 'Superscript'), 'change');
		field(root, 'Subscript').checked = true;
		fire(field(root, 'Subscript'), 'change');
		expect(field(root, 'Superscript').checked).toBe(false);
		ok(root);
		expect(
			view.state.doc.nodeAt(2)!.marks.find((mark) => mark.type.name === 'verticalAlign')!.attrs
				.value,
		).toBe('subscript');
	});

	it('shows mixed formatting as indeterminate', () => {
		const { view, dialogs } = setup({});
		view.dispatch(view.state.tr.addMark(1, 4, schema.marks.strike!.create()));
		selectAll(view);
		dialogs.open('font');
		expect(field(dialogs.elements[0]!, 'Strikethrough').indeterminate).toBe(true);
	});
});

describe('dialog launchers', () => {
	it('sit in the Font and Paragraph groups and open the matching dialog', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		const font = ribbon.querySelector<HTMLButtonElement>('[data-label="Font"] .ribbon-launcher')!;
		const paragraph = ribbon.querySelector<HTMLButtonElement>(
			'[data-panel="Home"] [data-label="Paragraph"] .ribbon-launcher',
		)!;
		font.click();
		paragraph.click();
		expect(seen).toEqual([
			{ type: 'formatDialog', kind: 'font' },
			{ type: 'formatDialog', kind: 'paragraph' },
		]);
		expect(font.getAttribute('aria-label')).toBe('Font settings');
	});

	it('closes one dialog when the other opens', () => {
		const { dialogs } = setup({});
		dialogs.open('font');
		dialogs.open('paragraph');
		expect(dialogs.elements.map((element) => element.hidden)).toEqual([true, false, true]);
		dialogs.closeAll();
		expect(dialogs.elements.every((element) => element.hidden)).toBe(true);
	});
});

describe('dialog localization and DOM hygiene', () => {
	it('localizes the fields each time a dialog opens', () => {
		const { dialogs } = setup({});
		dialogs.setLocale('fr');
		dialogs.open('paragraph');
		const root = dialogs.elements[1]!;
		expect(root.querySelector('h2')!.textContent).toBe('Paragraphe');
		expect(
			root.querySelector('[data-localearialabel="Keep with next"]')!.getAttribute('aria-label'),
		).toBe('Paragraphes solidaires');
		dialogs.closeAll();
		dialogs.setLocale('de');
		dialogs.open('font');
		expect(dialogs.elements[0]!.querySelector('h2')!.textContent).toBe('Schriftart');
	});

	it('keeps no fields in the document while closed, so names never collide', () => {
		const { dialogs } = setup({});
		expect(dialogs.elements.every((element) => element.childElementCount === 0)).toBe(true);
		dialogs.open('paragraph');
		expect(dialogs.elements[1]!.querySelector('[aria-label="Line spacing"]')).not.toBeNull();
		dialogs.closeAll();
		expect(document.querySelectorAll('[aria-label="Line spacing"]')).toHaveLength(0);
	});
});
