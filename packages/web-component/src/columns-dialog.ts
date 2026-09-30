import { twips, type SectionColumns, type SectionProperties } from '@christophervr/docx-core';
import { checkbox, dialogButton, fieldset, labelled, numberInput, row } from './dialog-fields';
import {
	changeColumnGap,
	changeColumnWidth,
	columnTextWidth,
	equalColumnDraft,
	readColumnDraft,
	validColumnDraft,
	type ColumnDraft,
} from './column-settings';
import type { FormatDialog } from './font-dialog';
import { localizeElement, translate, type EditorLocale } from './localization';

export interface ColumnsHost {
	section(): SectionProperties | undefined;
	canEdit(): boolean;
	apply(columns: SectionColumns): void;
	restoreFocus(): void;
}

/** Current-section column settings; width edits adjust the adjacent column as Word does. */
export function createColumnsDialog(host: ColumnsHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Columns');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Columns';
	const count = numberInput(1, 45, 1);
	const gap = numberInput(0, 22, 0.01);
	const width = document.createElement('output');
	const equal = checkbox('Equal column width');
	const separator = checkbox('Line between');
	const individual = document.createElement('div');
	individual.style.maxHeight = '240px';
	individual.style.overflow = 'auto';
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const content = [
		heading,
		fieldset(
			'Columns',
			labelled('Number of columns', count),
			row(labelled('Column width (inches)', width), labelled('Column spacing (inches)', gap)),
			equal.wrapper,
			separator.wrapper,
			individual,
		),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	let source: SectionProperties | undefined;
	let draft: ColumnDraft = { widths: [], gaps: [] };
	let controls: { width: HTMLInputElement; gap: HTMLInputElement; group: HTMLElement }[] = [];
	const total = () => (source ? columnTextWidth(source) : 0);
	const validCount = () =>
		count.value !== '' &&
		Number.isInteger(Number(count.value)) &&
		Number(count.value) >= 1 &&
		Number(count.value) <= 45;
	const validate = () => {
		gap.disabled = !equal.input.checked || Number(count.value) === 1;
		individual.hidden = equal.input.checked;
		const number = Number(count.value);
		const spacing = Math.round(Number(gap.value) * 1440);
		const size = (total() - spacing * (number - 1)) / number;
		width.value = equal.input.checked && Number.isFinite(size) ? (size / 1440).toFixed(3) : '—';
		const valid =
			validCount() &&
			(equal.input.checked
				? gap.value !== '' &&
					Number.isFinite(spacing) &&
					spacing >= 0 &&
					spacing <= 31680 &&
					size >= 720
				: draft.widths.length === number &&
					validColumnDraft(draft, total()) &&
					controls.every((control) => control.width.value !== '' && control.gap.value !== ''));
		message.textContent = valid
			? ''
			: 'Enter whole columns and spacing that leaves at least 0.5 inch for each column.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	const refreshControls = (except?: HTMLInputElement) => {
		controls.forEach((control, i) => {
			if (control.width !== except)
				control.width.value = String(Math.round((draft.widths[i]! / 1440) * 1000) / 1000);
			if (control.gap !== except)
				control.gap.value = String(Math.round((draft.gaps[i]! / 1440) * 1000) / 1000);
		});
	};
	const renderIndividual = () => {
		controls = draft.widths.map((_, i) => {
			const columnWidth = numberInput(0.5, 22, 0.001);
			const columnGap = numberInput(0, 22, 0.001);
			columnGap.disabled = i === draft.widths.length - 1;
			const title = document.createElement('strong');
			title.textContent = `${translate(locale, 'Column')} ${i + 1}`;
			const pair = row(
				labelled('Column width (inches)', columnWidth),
				labelled('Column spacing (inches)', columnGap),
			);
			const group = document.createElement('div');
			group.setAttribute('role', 'group');
			group.setAttribute('aria-label', title.textContent);
			group.append(title, pair);
			const edit = (input: HTMLInputElement, apply: (value: number) => void) => {
				if (input.value !== '' && Number.isFinite(Number(input.value))) {
					apply(Math.round(Number(input.value) * 1440));
					refreshControls(input);
				}
				validate();
			};
			columnWidth.addEventListener('input', () =>
				edit(columnWidth, (value) => changeColumnWidth(draft, i, value)),
			);
			columnGap.addEventListener('input', () =>
				edit(columnGap, (value) => changeColumnGap(draft, i, value)),
			);
			return { width: columnWidth, gap: columnGap, group };
		});
		individual.replaceChildren(...controls.map((control) => control.group));
		localizeElement(individual, locale);
		refreshControls();
	};
	count.addEventListener('input', () => {
		if (validCount()) {
			draft = equalColumnDraft(total(), Number(count.value), Math.round(Number(gap.value) * 1440));
			renderIndividual();
		}
		validate();
	});
	gap.addEventListener('input', validate);
	equal.input.addEventListener('change', () => {
		if (!equal.input.checked && validCount()) {
			draft = equalColumnDraft(total(), Number(count.value), Math.round(Number(gap.value) * 1440));
			renderIndividual();
		}
		validate();
	});
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (!source || !host.canEdit() || !validate()) return;
		const common = {
			count: Number(count.value),
			equalWidth: equal.input.checked,
			separator: separator.input.checked,
			spacingTwips: twips(Math.round(Number(gap.value) * 1440)),
		};
		host.apply(
			equal.input.checked
				? common
				: {
						...common,
						widths: draft.widths.map((value, i) => ({
							widthTwips: twips(value),
							...(i < draft.widths.length - 1 ? { spacingTwips: twips(draft.gaps[i]!) } : {}),
						})),
					},
		);
		hide();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			hide();
		} else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		close: hide,
		get isOpen() {
			return !element.hidden;
		},
		setLocale(next) {
			locale = next;
		},
		open() {
			source = host.section();
			if (!source || !host.canEdit()) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			count.value = String(source.columns.count);
			gap.value = String((source.columns.spacingTwips ?? 720) / 1440);
			equal.input.checked = source.columns.equalWidth;
			separator.input.checked = Boolean(source.columns.separator);
			draft = readColumnDraft(source);
			renderIndividual();
			validate();
			element.hidden = false;
			count.focus();
		},
	};
}
