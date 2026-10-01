// @vitest-environment jsdom
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { createBordersDialog } from './borders-dialog';
import { applyBordersAndShading, readBordersAndShading } from './paragraph-decoration';
import { schema } from './schema';

function editor() {
	const doc = schema.node('doc', null, [
		schema.nodes.paragraph!.create({ id: 'p' }, schema.text('text')),
	]);
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc, schema }),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2)));
	return view;
}

describe('Borders and Shading', () => {
	it('applies a pen to the chosen sides with a fill, and reads it back', () => {
		const view = editor();
		expect(
			applyBordersAndShading(view, {
				sides: { top: true, bottom: true },
				style: 'double',
				sizeEighthPoints: 12,
				color: '#ff0000',
				fill: '#ffee99',
			}),
		).toBe(true);
		const attrs = view.state.doc.firstChild!.attrs;
		expect(Object.keys(attrs.borders).sort()).toEqual(['bottom', 'top']);
		expect(attrs.borders.top).toMatchObject({
			style: 'double',
			sizeEighthPoints: 12,
			color: '#ff0000',
		});
		expect(attrs.shadingFill).toBe('#FFEE99');
		expect(readBordersAndShading(view.state)).toMatchObject({
			sides: { top: true, bottom: true, left: false, right: false },
			style: 'double',
			sizeEighthPoints: 12,
			fill: '#FFEE99',
		});
		applyBordersAndShading(view, {
			sides: {},
			style: 'single',
			sizeEighthPoints: 4,
			color: null,
			fill: null,
		});
		expect(view.state.doc.firstChild!.attrs.borders).toBeNull();
		expect(view.state.doc.firstChild!.attrs.shadingFill).toBeNull();
	});

	it('applies the dialog fields on OK', () => {
		const view = editor();
		const dialog = createBordersDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const box = (name: string) =>
			dialog.element.querySelector<HTMLInputElement>(`input[aria-label="${name}"]`)!;
		box('Left').checked = true;
		[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		expect(Object.keys(view.state.doc.firstChild!.attrs.borders)).toEqual(['left']);
		expect(dialog.isOpen).toBe(false);
		dialog.element.remove();
	});

	it('offers Apply to inside a table and draws cell edges instead of paragraph borders', () => {
		const cell = (id: string) =>
			schema.nodes.tableCell!.create(null, schema.nodes.paragraph!.create({ id }, schema.text(id)));
		const table = schema.nodes.table!.create({ id: 't' }, [
			schema.nodes.tableRow!.create(null, [cell('a'), cell('b')]),
		]);
		const doc = schema.node('doc', null, [table]);
		const view = new EditorView(document.createElement('div'), {
			state: EditorState.create({ doc, schema }),
		});
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 4)));
		const dialog = createBordersDialog(() => view);
		document.body.append(dialog.element);
		dialog.open();
		const field = (name: string) =>
			dialog.element.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
		expect(field('Apply to').value).toBe('cell');
		field('Right').checked = true;
		[...dialog.element.querySelectorAll('button')].find((b) => b.textContent === 'OK')!.click();
		const cells = view.state.doc.firstChild!.firstChild!;
		expect(JSON.parse(cells.child(0).attrs.directBorders).right.style).toBe('single');
		expect(JSON.parse(cells.child(1).attrs.directBorders).left.style).toBe('single');
		dialog.element.remove();
	});
});
