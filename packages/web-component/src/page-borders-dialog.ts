import type { PageBorders, SectionProperties } from 'docx-core';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
} from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';

export interface PageBordersHost {
	section(): SectionProperties | undefined;
	canEdit(): boolean;
	/** Replaces the current section's page borders; undefined removes them. */
	apply(borders: PageBorders | undefined): void;
	restoreFocus(): void;
}

const WIDTHS: ReadonlyArray<readonly [string, string]> = [
	['2', '¼ pt'],
	['4', '½ pt'],
	['6', '¾ pt'],
	['8', '1 pt'],
	['12', '1½ pt'],
	['18', '2¼ pt'],
	['24', '3 pt'],
	['36', '4½ pt'],
	['48', '6 pt'],
];
const SIDES = ['top', 'left', 'bottom', 'right'] as const;
const SIDE_LABELS = { top: 'Top', left: 'Left', bottom: 'Bottom', right: 'Right' } as const;
const asHex = (value: unknown): string | null =>
	typeof value === 'string' && /^#?[0-9a-f]{6}$/i.test(value) ? `#${value.replace('#', '')}` : null;

/**
 * Word's Page Borders for the section holding the selection: the sides, one pen, where the lines
 * are measured from, and which pages show them. They print in Print Layout; the continuous editing
 * surface has no pages to put them on. Art borders are not offered, and a pen chosen here replaces an
 * imported art side with a plain line.
 */
export function createPageBordersDialog(host: PageBordersHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-page-borders-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Page borders');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Page borders';
	const setting = selectOf([
		['none', 'None'],
		['box', 'Box'],
		['custom', 'Custom'],
	]);
	const sides = Object.fromEntries(
		SIDES.map((side) => [side, checkbox(SIDE_LABELS[side])]),
	) as Record<(typeof SIDES)[number], ReturnType<typeof checkbox>>;
	const style = selectOf([
		['single', 'Single'],
		['double', 'Double'],
		['dotted', 'Dotted'],
		['dashed', 'Dashed'],
		['thick', 'Thick'],
	]);
	const width = selectOf(WIDTHS);
	const color = document.createElement('input');
	color.type = 'color';
	const automatic = checkbox('Automatic');
	const from = selectOf([
		['page', 'Edge of page'],
		['text', 'Text'],
	]);
	const distance = numberInput(0, 31, 1);
	const display = selectOf([
		['allPages', 'All pages'],
		['firstPage', 'First page only'],
		['notFirstPage', 'All except first page'],
	]);
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const content = [
		heading,
		labelled('Setting', setting),
		fieldset('Borders', row(...SIDES.map((side) => sides[side].wrapper))),
		row(labelled('Style', style), labelled('Width', width)),
		row(labelled('Color', color), automatic.wrapper),
		row(
			labelled('Measure from', from),
			labelled('Distance (points)', distance),
			labelled('Show on', display),
		),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	let existing: PageBorders | undefined;

	const validate = () => {
		const off = setting.value === 'none';
		for (const control of [style, width, from, distance, display, automatic.input])
			control.disabled = off;
		for (const side of SIDES) sides[side].input.disabled = setting.value !== 'custom';
		color.disabled = off || automatic.input.checked;
		const valid =
			off ||
			(distance.value !== '' &&
				Number.isInteger(Number(distance.value)) &&
				Number(distance.value) >= 0 &&
				Number(distance.value) <= 31 &&
				SIDES.some((side) => sides[side].input.checked));
		message.textContent = valid
			? ''
			: 'Choose at least one side and a whole distance from 0 to 31 points.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	setting.addEventListener('change', () => {
		if (setting.value === 'box') for (const side of SIDES) sides[side].input.checked = true;
		validate();
	});
	for (const control of [style, width, color, automatic.input, from, distance, display])
		for (const type of ['input', 'change']) control.addEventListener(type, validate);
	for (const side of SIDES) sides[side].input.addEventListener('change', validate);

	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (!host.canEdit() || !validate()) return;
		if (setting.value === 'none') host.apply(undefined);
		else {
			const pen = {
				style: style.value as 'single',
				sizeEighthPoints: Number(width.value),
				spacePoints: Number(distance.value),
				...(automatic.input.checked ? {} : { color: color.value.toUpperCase() }),
			};
			const borders: PageBorders = {
				...Object.fromEntries(
					SIDES.filter((side) => sides[side].input.checked).map((side) => [side, pen]),
				),
				offsetFrom: from.value as 'page' | 'text',
				...(display.value !== 'allPages' ? { display: display.value as 'firstPage' } : {}),
				...(existing?.zOrder ? { zOrder: existing.zOrder } : {}),
			};
			host.apply(borders);
		}
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
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return !element.hidden;
		},
		open() {
			const section = host.section();
			if (!section || !host.canEdit()) return;
			existing = section.pageBorders;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const present = SIDES.filter((side) => existing?.[side]);
			setting.value = !present.length ? 'none' : present.length === 4 ? 'box' : 'custom';
			for (const side of SIDES) sides[side].input.checked = Boolean(existing?.[side]);
			const pen = existing?.top ?? existing?.left ?? existing?.bottom ?? existing?.right;
			const known = ['single', 'double', 'dotted', 'dashed', 'thick'];
			style.value = known.includes(pen?.style ?? '') ? (pen!.style as string) : 'single';
			width.value = WIDTHS.some(([value]) => value === String(pen?.sizeEighthPoints))
				? String(pen?.sizeEighthPoints)
				: '4';
			automatic.input.checked = !asHex(pen?.color);
			color.value = asHex(pen?.color) ?? '#000000';
			from.value = existing?.offsetFrom ?? 'page';
			distance.value = String(pen?.spacePoints ?? 24);
			display.value = existing?.display ?? 'allPages';
			validate();
			element.hidden = false;
			setting.focus();
		},
	};
}
