// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { modelToDoc } from './model-adapter';
import { createMultilevelListDialog } from './multilevel-list-dialog';

afterEach(() => document.body.replaceChildren());

function setup() {
	const model = createDocument();
	const view = new EditorView(document.createElement('div'), {
		state: EditorState.create({ doc: modelToDoc(model) }),
	});
	let allowed = true;
	const dialog = createMultilevelListDialog(
		() => view,
		() => model,
		() => allowed,
	);
	document.body.append(dialog.element);
	const field = (name: string) =>
		dialog.element.querySelector<HTMLInputElement & HTMLSelectElement>(`[aria-label="${name}"]`)!;
	const set = (name: string, value: string, event = 'input') => {
		field(name).value = value;
		field(name).dispatchEvent(new Event(event));
	};
	const button = (text: string) =>
		[...dialog.element.querySelectorAll<HTMLButtonElement>('button')].find(
			(button) => button.textContent === text,
		)!;
	return {
		model,
		view,
		dialog,
		field,
		set,
		button,
		allow: (value: boolean) => {
			allowed = value;
		},
	};
}

it('holds invalid level edits for correction and cancels without adding a definition', () => {
	const s = setup();
	s.dialog.open();
	s.set('Enter formatting for number', '%9.');
	expect(s.button('OK').disabled).toBe(true);
	s.set('Level to modify', '2', 'change');
	expect(s.field('Level to modify').value).toBe('0');
	s.set('Enter formatting for number', '%1.');
	s.set('Start at', '2.5');
	expect(s.button('OK').disabled).toBe(true);
	s.set('Start at', '3');
	expect(s.button('OK').disabled).toBe(false);
	s.button('Cancel').click();
	expect(s.model.numberingCatalog).toBeUndefined();
	expect(s.view.state.doc.firstChild!.attrs.numId).toBeNull();
	s.view.destroy();
});

it('rechecks editing permission at submission and localizes dialog controls', () => {
	const s = setup();
	s.allow(false);
	s.dialog.open();
	expect(s.dialog.isOpen).toBe(false);
	s.allow(true);
	s.dialog.open();
	s.allow(false);
	s.button('OK').click();
	expect(s.model.numberingCatalog).toBeUndefined();
	s.dialog.close();
	s.allow(true);
	s.dialog.setLocale('fr');
	s.dialog.open();
	expect(s.dialog.element.getAttribute('aria-label')).toBe(
		'Définir une nouvelle liste à plusieurs niveaux',
	);
	expect(s.dialog.element.querySelector('[aria-label="Niveau à modifier"]')).not.toBeNull();
	s.view.destroy();
});

it('stores marker font settings on the level and rejects an invalid colour', () => {
	const s = setup();
	s.dialog.open();
	s.set('Color (#rrggbb, blank for automatic)', 'blue');
	expect(s.button('OK').disabled).toBe(true);
	s.set('Color (#rrggbb, blank for automatic)', '#1F4E79');
	s.set('Font', 'Georgia');
	s.set('Size (pt)', '14');
	s.field('Bold').checked = true;
	s.field('Bold').dispatchEvent(new Event('input'));
	expect(s.button('OK').disabled).toBe(false);
	s.button('OK').click();
	const level = s.model.numberingCatalog!.abstractNums['0']!.levels[0]!;
	expect(level.markerFormat).toEqual({
		fontFamily: 'Georgia',
		fontSizeHalfPoints: 28,
		color: '#1f4e79',
		bold: true,
	});
	s.view.destroy();
});
