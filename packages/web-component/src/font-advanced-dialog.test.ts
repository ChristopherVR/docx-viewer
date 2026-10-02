// @vitest-environment jsdom
import { createDocument, halfPoints, signedTwips } from 'docx-core';
import { history, undo } from 'prosemirror-history';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { createFontDialog } from './font-dialog';
import { docToModel, modelToDoc } from './model-adapter';
import { readFontFormat } from './font-format';
import { runStylesPlugin } from './run-styles';

// These contracts inspect editor state; browser contracts verify real font measurements.
let canvasStub: { mockRestore(): void };
beforeAll(() => {
	canvasStub = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
});
afterAll(() => canvasStub.mockRestore());

function setup() {
	const model = createDocument();
	model.characterStyles = {
		docDefaults: {
			textScalePercent: 150,
			characterSpacingTwips: signedTwips(40),
			kerningHalfPoints: halfPoints(24),
			positionHalfPoints: halfPoints(6),
		},
		styles: {},
		warnings: [],
	};
	model.blocks = [
		{
			type: 'paragraph',
			id: 'a',
			runs: [
				{ text: 'one' },
				{ text: 'two', textScalePercent: 100, positionHalfPoints: halfPoints(-6) },
			],
		},
	];
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [runStylesPlugin(() => model), history()],
		}),
	});
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 7)));
	const dialog = createFontDialog(() => view);
	document.body.append(dialog.element);
	dialog.open();
	const root = dialog.element;
	const field = (name: string) =>
		root.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
	const fill = (name: string, value: string, event = 'input') => {
		const input = field(name);
		input.value = value;
		input.dispatchEvent(new Event(event, { bubbles: true }));
	};
	const ok = () =>
		[...root.querySelectorAll<HTMLButtonElement>('button')]
			.find((button) => button.textContent === 'OK')!
			.click();
	const dispose = () => {
		dialog.close();
		view.destroy();
		root.remove();
	};
	return { model, view, dialog, root, field, fill, ok, dispose };
}

describe('Advanced Font dialog', () => {
	it('leaves mixed ligatures untouched when another Advanced setting changes', () => {
		const s = setup();
		const type = s.view.state.schema.marks.runProperties!;
		s.view.dispatch(s.view.state.tr.addMark(1, 4, type.create({ props: { ligatures: 'all' } })));
		s.dialog.open();
		expect(s.field('Ligatures').selectedIndex).toBe(-1);
		s.fill('Scale (%)', '125');
		s.ok();
		expect(readFontFormat(s.view.state).ligatures).toBeNull();
		const p = docToModel(s.view.state.doc, s.model).blocks[0]!;
		if (p.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(p.runs.map((run) => run.ligatures)).toEqual(['all', undefined]);
		s.dispose();
	});
	it('applies ligature selection and undoes it without changing other fields', () => {
		const s = setup();
		expect(s.field('Ligatures').value).toBe('none');
		s.fill('Ligatures', 'standardContextual', 'change');
		s.ok();
		expect(readFontFormat(s.view.state)).toMatchObject({
			ligatures: 'standardContextual',
			scale: null,
			position: null,
			spacing: 2,
		});
		undo(s.view.state, s.view.dispatch);
		expect(readFontFormat(s.view.state).ligatures).toBe('none');
		s.dispose();
	});
	it('shows mixed values without overwriting untouched properties', () => {
		const s = setup();
		expect(s.field('Scale (%)').value).toBe('');
		expect(s.field('Position').selectedIndex).toBe(-1);
		expect(s.field('Points and above').value).toBe('12');
		s.fill('Scale (%)', '125');
		s.ok();
		expect(readFontFormat(s.view.state)).toMatchObject({
			scale: 125,
			position: null,
			kerning: 12,
			spacing: 2,
		});
		const paragraph = docToModel(s.view.state.doc, s.model).blocks[0]!;
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(paragraph.runs.map((run) => run.positionHalfPoints)).toEqual([undefined, -6]);
		s.dispose();
	});

	it('resets inherited spacing, scale, position and kerning explicitly in one undo step', () => {
		const s = setup();
		const original = s.view.state.doc;
		s.fill('Font style', 'boldItalic', 'change');
		s.fill('Scale (%)', '100');
		s.fill('Character spacing', 'normal', 'change');
		s.fill('Position', 'normal', 'change');
		const checkbox = s.field('Kerning for fonts');
		checkbox.checked = false;
		checkbox.dispatchEvent(new Event('change', { bubbles: true }));
		s.ok();
		expect(readFontFormat(s.view.state)).toMatchObject({
			bold: true,
			italic: true,
			scale: 100,
			position: 0,
			spacing: 0,
			kerning: 0,
		});
		expect(undo(s.view.state, s.view.dispatch)).toBe(true);
		expect(s.view.state.doc.eq(original)).toBe(true);
		s.dispose();
	});

	it('keeps the dialog open and model unchanged for invalid advanced input', () => {
		const s = setup();
		const original = s.view.state.doc;
		s.fill('Scale (%)', '601');
		s.ok();
		expect(s.dialog.isOpen).toBe(true);
		expect(s.view.state.doc.eq(original)).toBe(true);
		s.fill('Scale (%)', '125.5');
		s.ok();
		expect(s.dialog.isOpen).toBe(true);
		s.fill('Scale (%)', '125');
		s.ok();
		expect(s.dialog.isOpen).toBe(false);
		s.dispose();
	});

	it('supports keyboard tab switching and resets to Font on reopening', () => {
		const s = setup();
		const tabs = s.root.querySelectorAll<HTMLElement>('[role=tab]');
		tabs[0]!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
		expect(tabs[1]!.getAttribute('aria-selected')).toBe('true');
		expect(document.activeElement).toBe(tabs[1]);
		s.dialog.close();
		s.dialog.open();
		expect(tabs[0]!.getAttribute('aria-selected')).toBe('true');
		s.dispose();
	});

	it('localizes the tab and controls', () => {
		const s = setup();
		s.dialog.close();
		s.dialog.setLocale('fr');
		s.dialog.open();
		expect(s.root.querySelector('[aria-label="Échelle (%)"]')).not.toBeNull();
		expect(s.root.querySelector('option[value="standardContextual"]')!.textContent).toBe(
			'Standard Contextuelles',
		);
		expect([...s.root.querySelectorAll('[role=tab]')].map((tab) => tab.textContent)).toContain(
			'Paramètres avancés',
		);
		s.dispose();
	});
});
