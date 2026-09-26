import type { Mark } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { closeHistory } from 'prosemirror-history';
import { isValidLanguageTag } from '@christophervr/docx-core';
import { schema } from './schema';

export type LanguageField = 'language' | 'eastAsiaLanguage' | 'bidiLanguage';
export type MultilingualAction =
	| { type: 'paragraphDirection'; value: 'inherit' | 'ltr' | 'rtl' }
	| { type: 'language'; key: LanguageField; value: string }
	| { type: 'runRtl'; value: 'inherit' | 'on' | 'off' };

const languageOptions: Array<[string, string]> = [
	['', 'Inherit'],
	['en-US', 'English (United States)'],
	['en-GB', 'English (United Kingdom)'],
	['ar-SA', 'Arabic (Saudi Arabia)'],
	['he-IL', 'Hebrew (Israel)'],
	['ja-JP', 'Japanese (Japan)'],
	['zh-CN', 'Chinese (Simplified)'],
	['zh-TW', 'Chinese (Traditional)'],
	['ko-KR', 'Korean (Korea)'],
	['fr-FR', 'French (France)'],
	['de-DE', 'German (Germany)'],
	['es-ES', 'Spanish (Spain)'],
	['hi-IN', 'Hindi (India)'],
	['ru-RU', 'Russian (Russia)'],
];

function emit(control: HTMLElement, detail: MultilingualAction): void {
	control.dispatchEvent(
		new CustomEvent('ribbon-action', { bubbles: true, composed: true, detail }),
	);
}

function languageControl(label: string, key: LanguageField): HTMLElement {
	const wrapper = document.createElement('label');
	wrapper.className = 'multilingual-language-control';
	const visibleLabel = document.createElement('span');
	visibleLabel.textContent = label;
	const select = document.createElement('select');
	select.setAttribute('aria-label', label);
	for (const [value, text] of languageOptions) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = text;
		select.append(option);
	}
	select.addEventListener('change', () => {
		const input = wrapper.querySelector<HTMLInputElement>('input')!;
		input.value = '';
		input.setCustomValidity('');
		emit(wrapper, { type: 'language', key, value: select.value });
	});
	const custom = document.createElement('input');
	custom.type = 'text';
	custom.autocomplete = 'off';
	custom.spellcheck = false;
	custom.placeholder = 'Custom BCP 47 tag';
	custom.setAttribute('aria-label', `Custom ${label.toLowerCase()} tag`);
	custom.addEventListener('change', () => {
		const value = custom.value.trim();
		if (value && !isValidLanguageTag(value)) {
			custom.setCustomValidity('Enter a BCP 47 language tag, such as en-US or ar-SA.');
			custom.reportValidity();
			return;
		}
		custom.setCustomValidity('');
		select.value = '';
		emit(wrapper, { type: 'language', key, value });
	});
	wrapper.append(visibleLabel, select, custom);
	return wrapper;
}

export function createMultilingualControls(): HTMLElement {
	const group = document.createElement('div');
	group.className = 'ribbon-group multilingual-controls';
	const direction = document.createElement('select');
	direction.setAttribute('aria-label', 'Paragraph direction');
	for (const [value, label] of [
		['inherit', 'Direction: inherit'],
		['ltr', 'Left to right'],
		['rtl', 'Right to left'],
	]) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = label;
		direction.append(option);
	}
	direction.addEventListener('change', () =>
		emit(group, {
			type: 'paragraphDirection',
			value: direction.value as 'inherit' | 'ltr' | 'rtl',
		}),
	);
	const runDirection = document.createElement('select');
	runDirection.setAttribute('aria-label', 'Run direction');
	for (const [value, label] of [
		['inherit', 'Run direction: inherit'],
		['on', 'Run right to left'],
		['off', 'Run explicit non-RTL'],
	]) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = label;
		runDirection.append(option);
	}
	runDirection.addEventListener('change', () =>
		emit(group, { type: 'runRtl', value: runDirection.value as 'inherit' | 'on' | 'off' }),
	);
	group.append(
		direction,
		languageControl('Text language', 'language'),
		languageControl('East Asian language', 'eastAsiaLanguage'),
		languageControl('Complex script language', 'bidiLanguage'),
		runDirection,
	);
	const note = document.createElement('span');
	note.className = 'multilingual-note';
	note.textContent = 'Language tags annotate text; they do not translate or spell-check it.';
	group.append(note);
	return group;
}

function selectedTextMarks(state: EditorState): Array<readonly Mark[]> {
	if (state.selection.empty) return [state.storedMarks || state.selection.$from.marks()];
	const marks: Array<readonly Mark[]> = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
		if (node.isText) marks.push(node.marks);
	});
	return marks.length ? marks : [state.storedMarks || state.selection.$from.marks()];
}

function commonValue(values: Array<string | null>): string | null | 'mixed' {
	const distinct = [...new Set(values)];
	return distinct.length === 1 ? distinct[0] : 'mixed';
}

function setSelectValue(select: HTMLSelectElement | null, value: string): void {
	if (!select) return;
	let option = [...select.options].find((entry) => entry.value === value);
	if (!option) {
		option = document.createElement('option');
		option.value = value;
		option.textContent = value === 'mixed' ? 'Mixed' : value;
		select.append(option);
	}
	select.value = value;
}

export function syncMultilingualControls(toolbar: HTMLElement, state: EditorState): void {
	const paragraphs: string[] = [];
	if (state.selection.empty)
		paragraphs.push(String(state.selection.$from.parent.attrs.direction ?? 'inherit'));
	else
		state.doc.nodesBetween(state.selection.from, state.selection.to, (node) => {
			if (node.type.name === 'paragraph')
				paragraphs.push(String(node.attrs.direction ?? 'inherit'));
		});
	setSelectValue(
		toolbar.querySelector<HTMLSelectElement>('[aria-label="Paragraph direction"]'),
		commonValue(paragraphs.length ? paragraphs : ['inherit'])!,
	);
	const marks = selectedTextMarks(state);
	const languageKeys: Array<[LanguageField, string]> = [
		['language', 'Text language'],
		['eastAsiaLanguage', 'East Asian language'],
		['bidiLanguage', 'Complex script language'],
	];
	for (const [key, label] of languageKeys) {
		const values = marks.map((items) => {
			const mark = items.find((item) => item.type === schema.marks.language);
			return (mark?.attrs[key] as string | null | undefined) ?? null;
		});
		const value = commonValue(values);
		setSelectValue(
			toolbar.querySelector<HTMLSelectElement>(`select[aria-label="${label}"]`),
			value === 'mixed' ? 'mixed' : (value ?? ''),
		);
		const input = toolbar.querySelector<HTMLInputElement>(
			`[aria-label="Custom ${label.toLowerCase()} tag"]`,
		);
		if (input) input.value = value && value !== 'mixed' ? value : '';
	}
	const rtlValues = marks.map((items) => {
		const mark = items.find((item) => item.type === schema.marks.runRtl);
		return mark ? (mark.attrs.value ? 'on' : 'off') : 'inherit';
	});
	const rtl = commonValue(rtlValues);
	setSelectValue(
		toolbar.querySelector<HTMLSelectElement>('[aria-label="Run direction"]'),
		rtl ?? 'inherit',
	);
}

function languageMarkAttrs(mark: Mark | undefined, key: LanguageField, value: string) {
	const attrs = {
		language: mark?.attrs.language ?? null,
		eastAsiaLanguage: mark?.attrs.eastAsiaLanguage ?? null,
		bidiLanguage: mark?.attrs.bidiLanguage ?? null,
	};
	attrs[key] = value || null;
	return attrs;
}

function applyLanguage(view: EditorView, key: LanguageField, value: string): void {
	const tag = value.trim();
	if (tag && !isValidLanguageTag(tag)) return;
	const { state } = view;
	const type = schema.marks.language;
	if (state.selection.empty) {
		const marks = state.storedMarks || state.selection.$from.marks();
		const existing = marks.find((mark) => mark.type === type);
		const attrs = languageMarkAttrs(existing, key, tag);
		const next = Object.values(attrs).some(Boolean) ? type.create(attrs) : undefined;
		view.dispatch(
			state.tr.setStoredMarks([
				...marks.filter((mark) => mark.type !== type),
				...(next ? [next] : []),
			]),
		);
		return;
	}
	let transaction = state.tr;
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (!node.isText) return;
		const start = Math.max(pos, state.selection.from);
		const end = Math.min(pos + node.nodeSize, state.selection.to);
		const existing = node.marks.find((mark) => mark.type === type);
		const attrs = languageMarkAttrs(existing, key, tag);
		const next = Object.values(attrs).some(Boolean) ? type.create(attrs) : undefined;
		transaction = transaction.removeMark(start, end, type);
		if (next) transaction = transaction.addMark(start, end, next);
	});
	view.dispatch(closeHistory(transaction));
}

function applyRunRtl(view: EditorView, value: 'inherit' | 'on' | 'off'): void {
	const { state } = view;
	const type = schema.marks.runRtl;
	const mark = value === 'inherit' ? undefined : type.create({ value: value === 'on' });
	if (state.selection.empty) {
		const marks = state.storedMarks || state.selection.$from.marks();
		view.dispatch(
			state.tr.setStoredMarks([
				...marks.filter((item) => item.type !== type),
				...(mark ? [mark] : []),
			]),
		);
		return;
	}
	const transaction = state.tr.removeMark(state.selection.from, state.selection.to, type);
	view.dispatch(
		closeHistory(
			mark ? transaction.addMark(state.selection.from, state.selection.to, mark) : transaction,
		),
	);
}

function applyDirection(view: EditorView, value: 'inherit' | 'ltr' | 'rtl'): void {
	const { state } = view;
	const positions: number[] = [];
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (node.type.name === 'paragraph') positions.push(pos);
	});
	if (!positions.length && state.selection.$from.parent.type.name === 'paragraph')
		positions.push(state.selection.$from.before());
	let transaction = state.tr;
	for (const pos of positions) {
		const node = transaction.doc.nodeAt(pos);
		if (node)
			transaction = transaction.setNodeMarkup(pos, undefined, {
				...node.attrs,
				direction: value === 'inherit' ? null : value,
			});
	}
	if (transaction.docChanged) view.dispatch(closeHistory(transaction).scrollIntoView());
}

export function applyMultilingualAction(view: EditorView, action: MultilingualAction): void {
	if (!view.editable) return;
	if (action.type === 'language') applyLanguage(view, action.key, action.value);
	else if (action.type === 'runRtl') applyRunRtl(view, action.value);
	else applyDirection(view, action.value);
}
