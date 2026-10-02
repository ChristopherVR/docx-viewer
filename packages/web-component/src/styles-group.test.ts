// @vitest-environment jsdom
import { createDocument } from 'docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { createRibbon } from './ribbon';
import { schema } from './schema';
import { recommendedCharacterStyles, recommendedStyles } from './style-gallery';
import { resetStylePicker, syncStylePicker } from './styles-group';

afterEach(() => (document.body.innerHTML = ''));

function setup() {
	const model = createDocument();
	model.characterStyles!.styles.Emphasis = {
		id: 'Emphasis',
		type: 'character',
		name: 'Emphasis',
		formatting: { italic: true },
	};
	const frame = document.createElement('div');
	const ribbon = createRibbon();
	const body = document.createElement('div');
	body.className = 'dve-body';
	frame.append(ribbon, body);
	document.body.append(frame);
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p' }, schema.text('Some text')),
	]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
	syncStylePicker(ribbon, view, model, 'en');
	return { ribbon, view, model, body };
}
const tiles = (root: ParentNode) =>
	[...root.querySelectorAll<HTMLButtonElement>('.style-tile')].map((tile) => tile.dataset.styleId);
const button = (root: ParentNode, label: string) =>
	root.querySelector<HTMLButtonElement>(`button[aria-label="${label}"]`)!;

describe('recommended styles', () => {
	it('keeps the default first and drops technical paragraph and character styles', () => {
		const all = ['Normal', 'Title', 'heading 1', 'toc 1', 'footnote text', 'Header', 'Quote'].map(
			(name) => ({ id: name, name }),
		);
		expect(recommendedStyles(all).map((s) => s.name)).toEqual([
			'Normal',
			'Title',
			'heading 1',
			'Quote',
		]);
		const chars = [
			'Emphasis',
			'Strong',
			'Hyperlink',
			'footnote Text Char',
			'endnote reference',
		].map((name) => ({ id: name, name }));
		expect(recommendedCharacterStyles(chars).map((s) => s.name)).toEqual(['Emphasis', 'Strong']);
	});
});

describe('Styles group', () => {
	it('shows recommended styles as tiles and marks the current one', () => {
		const { ribbon } = setup();
		const gallery = ribbon.querySelector('[data-style-gallery]')!;
		const ids = tiles(gallery);
		expect(ids[0]).toBe('');
		expect(ids).toContain('Heading1');
		expect(ids).not.toContain('TOC1');
		expect(gallery.querySelector('[aria-pressed="true"]')!.getAttribute('data-style-id')).toBe('');
	});

	it('has previous, next and more controls, and a launcher for the pane', () => {
		const { ribbon } = setup();
		const group = ribbon.querySelector('[data-styles-group]')!;
		for (const label of ['Previous styles', 'Next styles', 'More styles', 'Styles pane'])
			expect(button(group, label), label).not.toBeNull();
		expect(ribbon.querySelectorAll('[data-styles-group]')).toHaveLength(1);
	});

	it('applies a gallery tile to the paragraph', () => {
		const { ribbon, view } = setup();
		ribbon.querySelector<HTMLButtonElement>('.style-tile[data-style-id="Heading1"]')!.click();
		expect(view.state.doc.firstChild!.attrs.style).toBe('Heading1');
	});

	it('opens the expanded gallery with both kinds of style, then Clear Formatting and the pane', () => {
		const { ribbon, view } = setup();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		button(ribbon, 'More styles').click();
		const popover = document.querySelector('.styles-popover')!;
		expect(popover).not.toBeNull();
		expect(tiles(popover)).toContain('Heading2');
		expect(tiles(popover)).not.toContain('TOC1');
		expect(popover.textContent).toContain('Character styles');
		expect(tiles(popover)).toContain('Emphasis');
		expect(tiles(popover)).not.toContain('Hyperlink');
		popover.querySelector<HTMLButtonElement>('.style-tile[data-style-id="Heading2"]')!.click();
		expect(view.state.doc.firstChild!.attrs.style).toBe('Heading2');
		expect(document.querySelector('.styles-popover')).toBeNull();

		button(ribbon, 'More styles').click();
		[...document.querySelectorAll<HTMLButtonElement>('.styles-command')]
			.find((b) => b.textContent === 'Clear formatting')!
			.click();
		expect(seen).toContainEqual({ type: 'clear' });
	});

	it('toggles the Styles pane and lists every style in it', () => {
		const { ribbon, body } = setup();
		const pane = body.querySelector<HTMLElement>('.dve-styles-pane')!;
		expect(pane.hidden).toBe(true);
		const launcher = button(ribbon, 'Styles pane');
		launcher.click();
		expect(pane.hidden).toBe(false);
		expect(launcher.getAttribute('aria-pressed')).toBe('true');
		expect(tiles(pane)).toContain('TOC1');
		expect(tiles(pane)).toContain('Heading1');
		pane.querySelector<HTMLButtonElement>('.style-tile[data-style-id="Title"]')!.click();
		launcher.click();
		expect(pane.hidden).toBe(true);
		expect(launcher.getAttribute('aria-pressed')).toBe('false');
	});

	it('keeps the hidden state select in step with the selection and removes everything on reset', () => {
		const { ribbon, view, body } = setup();
		const select = ribbon.querySelector<HTMLSelectElement>('[data-paragraph-styles]')!;
		view.dispatch(view.state.tr.setNodeAttribute(0, 'style', 'Heading3'));
		syncStylePicker(ribbon, view, createDocument(), 'en');
		expect(select.value).toBe('Heading3');
		resetStylePicker(ribbon);
		expect(ribbon.querySelector('[data-styles-group]')).toBeNull();
		expect(body.querySelector('.dve-styles-pane')).toBeNull();
	});

	it('disables the tiles for a read-only view', () => {
		const { ribbon, view, model } = setup();
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		syncStylePicker(ribbon, locked, model, 'en');
		expect(
			[...ribbon.querySelectorAll<HTMLButtonElement>('.ribbon-gallery .style-tile')].every(
				(tile) => tile.disabled,
			),
		).toBe(true);
	});
});
