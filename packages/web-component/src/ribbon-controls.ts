import type { EditorState } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import { schema } from './schema';
import { lineSpacingLabel, lineSpacingValue } from './line-spacing';
import { findLocalizedControl } from './localization';
import { selectionScript } from './script-state';

/** Synchronize the inline-format toggle buttons with the current selection. */
export function syncFormatControls(toolbar: HTMLElement, state: EditorState) {
	const marked = new Map<string, Set<boolean>>();
	for (const key of ['bold', 'italic', 'underline', 'strike', 'verticalAlign'])
		marked.set(key, new Set());
	const addMarks = (marks: readonly Mark[]) => {
		for (const key of ['bold', 'italic', 'underline', 'strike'])
			marked.get(key)!.add(marks.some((mark) => mark.type.name === key));
		const align = marks.find((mark) => mark.type.name === 'verticalAlign');
		marked.get('verticalAlign')!.add(Boolean(align));
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
		findLocalizedControl(toolbar, label)?.setAttribute(
			'aria-pressed',
			String(values.size === 1 && values.has(true)),
		);
	}
	const superButton = findLocalizedControl(toolbar, 'Superscript');
	const subButton = findLocalizedControl(toolbar, 'Subscript');
	const align = selectionScript(state);
	superButton?.setAttribute('aria-pressed', String(align === 'superscript'));
	subButton?.setAttribute('aria-pressed', String(align === 'subscript'));
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
		const control = findLocalizedControl<HTMLSelectElement>(toolbar, label);
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
		if (value === undefined) continue;
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
