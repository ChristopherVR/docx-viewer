import type { EditorView } from 'prosemirror-view';
import type { DocumentModel } from '@christophervr/docx-core';
import { translate, type EditorLocale } from './localization';
import { schema } from './schema';

/** The character style (`w:rStyle`) at the start of the selection, or '' when none applies. */
export function characterStyleAtSelection(view: EditorView): string {
	const { $from, from, empty } = view.state.selection;
	const marks = empty
		? (view.state.storedMarks ?? $from.marks())
		: (view.state.doc.nodeAt(from)?.marks ?? $from.marks());
	const mark = marks.find((item) => item.type === schema.marks.characterStyle);
	return mark ? String(mark.attrs.id) : '';
}

/** Applies (or with '' removes) a character style on the selection, or for the next typed text. */
export function applyCharacterStyle(view: EditorView, styleId: string): boolean {
	const type = schema.marks.characterStyle;
	const { from, to, empty } = view.state.selection;
	const tr = view.state.tr;
	if (empty) {
		tr.removeStoredMark(type);
		if (styleId) tr.addStoredMark(type.create({ id: styleId }));
	} else {
		tr.removeMark(from, to, type);
		if (styleId) tr.addMark(from, to, type.create({ id: styleId }));
	}
	view.dispatch(tr.scrollIntoView());
	return true;
}

/**
 * Keeps a "Character style" select beside the paragraph style picker. Changes are raised as a
 * `characterStyle` ribbon action so the editor applies them to its current view.
 */
export function syncCharacterStylePicker(
	toolbar: HTMLElement,
	view: EditorView,
	model: DocumentModel,
	locale: EditorLocale,
) {
	const group = toolbar.querySelector('[data-paragraph-styles]')?.parentElement;
	if (!group) return;
	let select = group.querySelector<HTMLSelectElement>('[data-character-styles]');
	if (!select) {
		select = document.createElement('select');
		select.dataset.characterStyles = '';
		select.dataset.action = 'select';
		select.addEventListener('change', () =>
			select!.dispatchEvent(
				new CustomEvent('ribbon-action', {
					bubbles: true,
					composed: true,
					detail: { type: 'characterStyle', value: select!.value },
				}),
			),
		);
		group.append(select);
	}
	select.setAttribute('aria-label', translate(locale, 'Character style'));
	const definitions = Object.values(model.characterStyles?.styles ?? {}).filter(
		(style) => style.type === 'character',
	);
	const selected = characterStyleAtSelection(view);
	const options = [
		{ id: '', name: translate(locale, 'No character style') },
		...definitions.map((style) => ({ id: style.id, name: style.name || style.id })),
	];
	if (selected && !definitions.some((style) => style.id === selected))
		options.push({ id: selected, name: selected });
	select.replaceChildren(...options.map((style) => new Option(style.name, style.id)));
	select.value = selected;
	select.disabled = !view.editable || !definitions.length;
}
