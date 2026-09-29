// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { setBorders, setShading } from './paragraph-decoration';
import { createRibbon } from './ribbon';
import { schema } from './schema';

function editor(...attrs: Array<Record<string, unknown>>) {
	const doc = schema.node(
		'doc',
		null,
		attrs.map((extra, index) =>
			schema.nodes.paragraph!.create({ id: `p${index}`, ...extra }, schema.text(`para ${index}`)),
		),
	);
	return new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
}
const selectAll = (view: EditorView) =>
	view.dispatch(
		view.state.tr.setSelection(
			TextSelection.create(view.state.doc, 1, view.state.doc.content.size - 1),
		),
	);
const borders = (view: EditorView, index = 0) =>
	view.state.doc.child(index).attrs.borders as Record<string, { style: string }> | null;

describe('setShading', () => {
	it('shades every selected paragraph and clears with null', () => {
		const view = editor({}, {});
		selectAll(view);
		expect(setShading(view, '#c0e6f5')).toBe(true);
		view.state.doc.forEach((p) => expect(p.attrs.shadingFill).toBe('#C0E6F5'));
		setShading(view, null);
		view.state.doc.forEach((p) => expect(p.attrs.shadingFill).toBeNull());
	});

	it('rejects anything that is not a hex colour, and read-only views', () => {
		const view = editor({});
		expect(setShading(view, 'yellow')).toBe(false);
		expect(setShading(view, '#fff')).toBe(false);
		const locked = new EditorView(document.createElement('div'), {
			state: view.state,
			editable: () => false,
		});
		expect(setShading(locked, '#ffffff')).toBe(false);
	});
});

describe('setBorders', () => {
	it("adds a bottom border with Word's default pen", () => {
		const view = editor({});
		setBorders(view, 'bottom');
		expect(view.state.doc.firstChild!.attrs.borders).toEqual({
			bottom: { style: 'single', sizeEighthPoints: 4, spacePoints: 1 },
		});
	});

	it('toggles a side off when every selected paragraph already has it', () => {
		const view = editor({}, {});
		selectAll(view);
		setBorders(view, 'top');
		expect(borders(view, 0)!.top).toBeDefined();
		setBorders(view, 'top');
		expect(borders(view, 0)).toBeNull();
		expect(borders(view, 1)).toBeNull();
	});

	it('adds to paragraphs that lack the side when only some have it', () => {
		const view = editor({ borders: { top: { style: 'single', sizeEighthPoints: 4 } } }, {});
		selectAll(view);
		setBorders(view, 'top');
		expect(borders(view, 1)!.top).toBeDefined();
		expect(borders(view, 0)!.top).toBeDefined();
	});

	it('draws all four sides for a box, adds the inside line, and keeps other sides', () => {
		const view = editor({ borders: { bottom: { style: 'single', sizeEighthPoints: 12 } } });
		setBorders(view, 'all');
		expect(Object.keys(borders(view)!).sort()).toEqual(['bottom', 'left', 'right', 'top']);
		setBorders(view, 'insideH');
		expect(Object.keys(borders(view)!)).toContain('between');
	});

	it('removes everything with none', () => {
		const view = editor({ borders: { top: { style: 'single', sizeEighthPoints: 4 } } });
		setBorders(view, 'none');
		expect(borders(view)).toBeNull();
	});
});

describe('shading and border controls', () => {
	it('emit actions; the colour button remembers its last colour', () => {
		const ribbon = createRibbon();
		const seen: unknown[] = [];
		ribbon.addEventListener('ribbon-action', (event) => seen.push((event as CustomEvent).detail));
		ribbon.querySelector<HTMLButtonElement>('button[aria-label="Shading"]')!.click();
		const menu = ribbon.querySelector<HTMLSelectElement>('select[aria-label="Borders"]')!;
		menu.value = 'outside';
		menu.dispatchEvent(new Event('change'));
		expect(seen).toEqual([
			{ type: 'shading', value: '#ffff00' },
			{ type: 'borders', preset: 'outside' },
		]);
	});
});
