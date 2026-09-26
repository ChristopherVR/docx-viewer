import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import { schema } from './schema';
import { lineSpacingLabel, lineSpacingValue } from './line-spacing';

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

/** Synchronize inline-format buttons and the highlight picker with the current selection. */
export function syncFormatControls(toolbar: HTMLElement, state: EditorState) {
	const marked = new Map<string, Set<boolean>>();
	for (const key of ['bold', 'italic', 'underline', 'strike', 'verticalAlign'])
		marked.set(key, new Set());
	const highlights = new Set<string>();
	const addMarks = (marks: readonly Mark[]) => {
		for (const key of ['bold', 'italic', 'underline', 'strike'])
			marked.get(key)!.add(marks.some((mark) => mark.type.name === key));
		const align = marks.find((mark) => mark.type.name === 'verticalAlign');
		marked.get('verticalAlign')!.add(Boolean(align));
		highlights.add(marks.find((mark) => mark.type.name === 'highlight')?.attrs.color || 'none');
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
	const labels: Record<string, string> = {
		bold: 'Bold',
		italic: 'Italic',
		underline: 'Underline',
		strike: 'Strikethrough',
	};
	for (const [key, label] of Object.entries(labels)) {
		const values = marked.get(key)!;
		toolbar
			.querySelector(`[aria-label="${label}"]`)
			?.setAttribute('aria-pressed', String(values.size === 1 && values.has(true)));
	}
	const superButton = toolbar.querySelector('[aria-label="Superscript"]');
	const subButton = toolbar.querySelector('[aria-label="Subscript"]');
	const aligns = new Set<string>();
	if (state.selection.empty) {
		const align = (state.storedMarks || state.selection.$from.marks()).find(
			(mark) => mark.type.name === 'verticalAlign',
		);
		aligns.add(align?.attrs.value || 'none');
	} else {
		let found = false;
		state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
			if (!node.isText) return;
			found = true;
			aligns.add(
				node.marks.find((mark) => mark.type.name === 'verticalAlign')?.attrs.value || 'none',
			);
		});
		if (!found) aligns.add('none');
	}
	superButton?.setAttribute('aria-pressed', String(aligns.size === 1 && aligns.has('superscript')));
	subButton?.setAttribute('aria-pressed', String(aligns.size === 1 && aligns.has('subscript')));
	const highlight = toolbar.querySelector<HTMLSelectElement>('[aria-label="Text highlight"]');
	if (highlight && highlights.size === 1) highlight.value = [...highlights][0];
	else if (highlight) highlight.selectedIndex = -1;
}

export function syncParagraphControls(toolbar: HTMLElement, state: EditorState) {
	const values: Record<'Spacing before' | 'Spacing after' | 'Line spacing', Set<string>> = {
		'Spacing before': new Set(),
		'Spacing after': new Set(),
		'Line spacing': new Set(),
	};
	const add = (attrs: Record<string, unknown>) => {
		values['Spacing before'].add(String(attrs.spacingBeforeTwips ?? 'inherited'));
		values['Spacing after'].add(String(attrs.spacingAfterTwips ?? 'inherited'));
		values['Line spacing'].add(lineSpacingValue(attrs));
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
		if (label === 'Line spacing') {
			for (const value of distinct) {
				if ([...control.options].some((option) => option.value === value)) continue;
				const option = document.createElement('option');
				option.value = value;
				option.textContent = lineSpacingLabel(value);
				control.append(option);
			}
		}
		if (distinct.length !== 1) {
			let mixed = control.querySelector<HTMLOptionElement>('option[value="mixed"]');
			if (!mixed) {
				mixed = document.createElement('option');
				mixed.value = 'mixed';
				mixed.textContent = label === 'Line spacing' ? 'Mixed paragraphs' : 'Mixed';
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
		if (label === 'Line spacing') {
			let option = [...control.options].find((item) => item.value === value);
			if (!option) {
				option = document.createElement('option');
				option.value = value;
				option.textContent = lineSpacingLabel(value);
				control.append(option);
			}
			control.value = value;
			continue;
		}
		if (![...control.options].some((option) => option.value === value)) {
			const option = document.createElement('option');
			option.value = value;
			option.textContent = `${(Number(value) / 20).toLocaleString()} pt`;
			control.append(option);
		}
		control.value = value;
	}
}
