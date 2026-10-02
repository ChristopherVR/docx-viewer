import type { Mark } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { isValidLanguageTag } from 'docx-core';
import { emit as emitEvent } from './events';
import { schema } from './schema';
import {
	findLocalizedControl,
	localeOf,
	localizeElement,
	normalizeEditorLocale,
	translate,
} from './localization';

import type { LanguageField, MultilingualAction } from './multilingual-commands';

export type { LanguageField, MultilingualAction };
export { applyMultilingualAction } from './multilingual-commands';

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
	emitEvent(control, 'ribbon-action', detail);
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
		if (value) option.dataset.languageTag = value;
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
			custom.setCustomValidity(
				translate(localeOf(custom), 'Enter a BCP 47 language tag, such as en-US or ar-SA.'),
			);
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

export function createMultilingualControls(locale = 'en'): HTMLElement {
	const group = document.createElement('div');
	group.className = 'ribbon-group multilingual-controls';
	const direction = document.createElement('select');
	direction.setAttribute('aria-label', 'Paragraph direction');
	const directionOptions: [string, string][] = [
		['inherit', 'Direction: inherit'],
		['ltr', 'Left to right'],
		['rtl', 'Right to left'],
	];
	for (const [value, label] of directionOptions) {
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
	const runDirectionOptions: [string, string][] = [
		['inherit', 'Run direction: inherit'],
		['on', 'Run right to left'],
		['off', 'Run explicit non-RTL'],
	];
	for (const [value, label] of runDirectionOptions) {
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
	localizeElement(group, normalizeEditorLocale(locale));
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
	return distinct.length === 1 ? (distinct[0] ?? null) : 'mixed';
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
		findLocalizedControl<HTMLSelectElement>(toolbar, 'Paragraph direction'),
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
			findLocalizedControl<HTMLSelectElement>(toolbar, label),
			value === 'mixed' ? 'mixed' : (value ?? ''),
		);
		const input = findLocalizedControl<HTMLInputElement>(
			toolbar,
			`Custom ${label.toLowerCase()} tag`,
		);
		if (input) input.value = value && value !== 'mixed' ? value : '';
	}
	const rtlValues = marks.map((items) => {
		const mark = items.find((item) => item.type === schema.marks.runRtl);
		return mark ? (mark.attrs.value ? 'on' : 'off') : 'inherit';
	});
	const rtl = commonValue(rtlValues);
	setSelectValue(
		findLocalizedControl<HTMLSelectElement>(toolbar, 'Run direction'),
		rtl ?? 'inherit',
	);
}
