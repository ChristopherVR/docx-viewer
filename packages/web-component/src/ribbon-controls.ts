import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import { schema } from './schema';

export function syncFontControls(toolbar: HTMLElement, state: EditorState) {
	const type = schema.marks.font;
	const defaults = { family: 'Calibri', size: 11, color: '#000000' };
	const values: Record<'family' | 'size' | 'color', Array<string | number>> = {
		family: [],
		size: [],
		color: [],
	};
	const addMarks = (marks: readonly Mark[]) => {
		const font = marks.find((mark) => mark.type === type);
		for (const key of Object.keys(values) as Array<keyof typeof values>)
			values[key].push(font?.attrs[key] ?? defaults[key]);
	};
	if (state.selection.empty) addMarks(state.storedMarks || state.selection.$from.marks());
	else {
		let found = false;
		state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
			if (!node.isText) return;
			found = true;
			addMarks(node.marks);
		});
		if (!found) addMarks(state.selection.$from.marks());
	}
	for (const key of Object.keys(values) as Array<keyof typeof values>) {
		const label = key === 'family' ? 'Font family' : key === 'size' ? 'Font size' : 'Font color';
		const select = toolbar.querySelector<HTMLSelectElement>(`[aria-label="${label}"]`);
		if (!select) continue;
		const distinct = [...new Set(values[key].map(String))];
		if (distinct.length !== 1) {
			select.selectedIndex = -1;
			continue;
		}
		const value = distinct[0];
		if (![...select.options].some((option) => option.value === value)) {
			const option = document.createElement('option');
			option.value = value;
			option.textContent = value;
			select.append(option);
		}
		select.value = value;
	}
}

export function syncParagraphControls(toolbar: HTMLElement, state: EditorState) {
	const values: Record<'Spacing before' | 'Spacing after', Set<string>> = {
		'Spacing before': new Set(),
		'Spacing after': new Set(),
	};
	const add = (attrs: Record<string, unknown>) => {
		values['Spacing before'].add(String(attrs.spacingBeforeTwips ?? 'inherited'));
		values['Spacing after'].add(String(attrs.spacingAfterTwips ?? 'inherited'));
	};
	if (state.selection.empty) add(state.selection.$from.parent.attrs);
	else {
		state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
			if (node.type.name === 'paragraph') add(node.attrs);
		});
	}
	for (const label of Object.keys(values) as Array<keyof typeof values>) {
		const control = toolbar.querySelector<HTMLSelectElement>(`[aria-label="${label}"]`);
		if (!control) continue;
		const distinct = [...values[label]];
		if (distinct.length !== 1) {
			let mixed = control.querySelector<HTMLOptionElement>('option[value="mixed"]');
			if (!mixed) {
				mixed = document.createElement('option');
				mixed.value = 'mixed';
				mixed.textContent = 'Mixed';
				control.append(mixed);
			}
			control.value = 'mixed';
			continue;
		}
		if (distinct[0] === 'inherited') {
			control.value = 'inherit';
			continue;
		}
		const value = distinct[0];
		if (![...control.options].some((option) => option.value === value)) {
			const option = document.createElement('option');
			option.value = value;
			option.textContent = `${(Number(value) / 20).toLocaleString()} pt`;
			control.append(option);
		}
		control.value = value;
	}
}
