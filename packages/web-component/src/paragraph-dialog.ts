import type { DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	setTriState,
} from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { focusView } from './focus-view';
import { localizeElement, type EditorLocale } from './localization';
import {
	applyParagraphFormat,
	readParagraphFormat,
	type Alignment,
	type LineRule,
	type ParagraphFormat,
	type Special,
} from './paragraph-format';

type BooleanField =
	| 'contextualSpacing'
	| 'widowControl'
	| 'keepNext'
	| 'keepLines'
	| 'pageBreakBefore'
	| 'suppressLineNumbers';

/** Word's Paragraph dialog: alignment, indentation, spacing, line spacing and pagination. */
export function createParagraphDialog(
	getView: () => EditorView | undefined,
	getModel: () => DocumentModel,
): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-paragraph-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Paragraph');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Paragraph';

	const align = selectOf([
		['left', 'Left'],
		['center', 'Centered'],
		['right', 'Right'],
		['justify', 'Justified'],
	]);
	const left = numberInput(0, 22, 0.05);
	const right = numberInput(0, 22, 0.05);
	const special = selectOf([
		['none', '(none)'],
		['firstLine', 'First line'],
		['hanging', 'Hanging'],
	]);
	const by = numberInput(0, 22, 0.05);
	const before = numberInput(0, 1584, 1);
	const after = numberInput(0, 1584, 1);
	const lineRule = selectOf([
		['single', 'Single spacing'],
		['oneAndHalf', 'One and a half'],
		['double', 'Double spacing'],
		['atLeast', 'At least'],
		['exactly', 'Exactly'],
		['multiple', 'Multiple'],
	]);
	const lineAt = numberInput(0.1, 132, 0.05);
	const contextual = checkbox('Do not add space between paragraphs of the same style');
	const pagination = {
		widowControl: checkbox('Widow/Orphan control'),
		keepNext: checkbox('Keep with next'),
		keepLines: checkbox('Keep lines together'),
		pageBreakBefore: checkbox('Page break before'),
		suppressLineNumbers: checkbox('Suppress line numbers'),
	};

	const toggles: Array<[BooleanField, ReturnType<typeof checkbox>]> = [
		['contextualSpacing', contextual],
		['widowControl', pagination.widowControl],
		['keepNext', pagination.keepNext],
		['keepLines', pagination.keepLines],
		['pageBreakBefore', pagination.pageBreakBefore],
		['suppressLineNumbers', pagination.suppressLineNumbers],
	];

	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, confirm);
	element.append(
		heading,
		fieldset('General', labelled('Alignment', align)),
		fieldset(
			'Indentation',
			row(labelled('Left', left), labelled('Right', right)),
			row(labelled('Special', special), labelled('By', by)),
		),
		fieldset(
			'Spacing',
			row(labelled('Before', before), labelled('After', after)),
			row(labelled('Line spacing', lineRule), labelled('At', lineAt)),
			contextual.wrapper,
		),
		fieldset('Pagination', ...Object.values(pagination).map((item) => item.wrapper)),
		actions,
	);

	const dirty = new Set<keyof ParagraphFormat>();
	const mark = (...fields: Array<keyof ParagraphFormat>) => {
		for (const field of fields) dirty.add(field);
	};
	const syncEnabled = () => {
		by.disabled = special.value === 'none';
		lineAt.disabled = ['single', 'oneAndHalf', 'double'].includes(lineRule.value);
	};
	align.addEventListener('change', () => mark('align'));
	left.addEventListener('input', () => mark('leftInches'));
	right.addEventListener('input', () => mark('rightInches'));
	special.addEventListener('change', () => {
		if (special.value !== 'none' && !Number(by.value)) by.value = '0.5';
		mark('special', 'specialInches');
		syncEnabled();
	});
	by.addEventListener('input', () => mark('specialInches', 'special'));
	before.addEventListener('input', () => mark('beforePt'));
	after.addEventListener('input', () => mark('afterPt'));
	lineRule.addEventListener('change', () => {
		if (!lineAt.value || lineAt.disabled)
			lineAt.value = lineRule.value === 'multiple' ? '1.1' : '12';
		mark('lineRule', 'lineAt');
		syncEnabled();
	});
	lineAt.addEventListener('input', () => mark('lineAt', 'lineRule'));
	for (const [field, item] of toggles)
		item.input.addEventListener('change', () => {
			item.input.indeterminate = false;
			mark(field);
		});

	const number = (input: HTMLInputElement) => {
		const value = Number(input.value);
		return input.value !== '' && Number.isFinite(value) ? value : undefined;
	};
	const collect = (): Partial<ParagraphFormat> => {
		const out: Partial<Record<keyof ParagraphFormat, unknown>> = {};
		if (dirty.has('align')) out.align = align.value as Alignment;
		if (dirty.has('leftInches')) out.leftInches = number(left);
		if (dirty.has('rightInches')) out.rightInches = number(right);
		if (dirty.has('special') || dirty.has('specialInches')) {
			out.special = special.value as Special;
			out.specialInches = special.value === 'none' ? 0 : (number(by) ?? 0.5);
		}
		if (dirty.has('beforePt')) out.beforePt = number(before);
		if (dirty.has('afterPt')) out.afterPt = number(after);
		if (dirty.has('lineRule') || dirty.has('lineAt')) {
			out.lineRule = lineRule.value as LineRule;
			out.lineAt = number(lineAt);
		}
		for (const [field, item] of toggles) if (dirty.has(field)) out[field] = item.input.checked;
		for (const key of Object.keys(out) as Array<keyof ParagraphFormat>)
			if (out[key] === undefined) delete out[key];
		return out as Partial<ParagraphFormat>;
	};

	const close = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		if (view) applyParagraphFormat(view, collect());
		close();
	};
	cancel.addEventListener('click', close);
	confirm.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});

	const show = (input: HTMLInputElement, value: number | null) => {
		input.value = value === null ? '' : String(Math.round(value * 100) / 100);
	};
	// The fields exist in the document only while the dialog is open, so their names never
	// collide with the ribbon's controls of the same name (Line spacing, Font size, ...).
	const content = [...element.childNodes];
	element.replaceChildren();
	let locale: EditorLocale = 'en';
	return {
		element,
		open() {
			const view = getView();
			if (!view?.editable) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const format = readParagraphFormat(view.state, getModel());
			dirty.clear();
			if (format.align === null) align.selectedIndex = -1;
			else align.value = format.align;
			show(left, format.leftInches);
			show(right, format.rightInches);
			if (format.special === null) special.selectedIndex = -1;
			else special.value = format.special;
			show(by, format.special === 'none' ? null : format.specialInches);
			show(before, format.beforePt);
			show(after, format.afterPt);
			if (format.lineRule === null) lineRule.selectedIndex = -1;
			else lineRule.value = format.lineRule;
			show(
				lineAt,
				['single', 'oneAndHalf', 'double'].includes(String(format.lineRule)) ? null : format.lineAt,
			);
			for (const [field, item] of toggles) setTriState(item.input, format[field]);
			syncEnabled();
			element.hidden = false;
			align.focus();
		},
		close,
		setLocale(next: EditorLocale) {
			locale = next;
		},
		get isOpen() {
			return !element.hidden;
		},
	};
}
