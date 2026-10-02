import { twips, type SectionProperties } from 'docx-core';
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

export type LineNumberSettings = NonNullable<SectionProperties['lineNumberSettings']>;
export interface LineNumberHost {
	section(): SectionProperties | undefined;
	canEdit(): boolean;
	apply(settings: LineNumberSettings | undefined): void;
	restoreFocus(): void;
}

/** Word's Line Numbers options, applied to the section holding the selection. */
export function createLineNumberDialog(host: LineNumberHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Line numbers');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Line numbers';
	const enabled = checkbox('Add line numbering');
	const start = numberInput(1, 32767, 1);
	const count = numberInput(1, 32767, 1);
	const automatic = checkbox('Automatic distance from text');
	const distance = numberInput(0, 22, 0.01);
	const restart = selectOf([
		['continuous', 'Continuous'],
		['newPage', 'Restart Each Page'],
		['newSection', 'Restart Each Section'],
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
		enabled.wrapper,
		fieldset(
			'Line numbers',
			row(labelled('Start at', start), labelled('Count by', count)),
			automatic.wrapper,
			row(labelled('Distance from text (inches)', distance), labelled('Numbering', restart)),
		),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	const validate = () => {
		for (const control of [start, count, automatic.input, restart])
			control.disabled = !enabled.input.checked;
		distance.disabled = !enabled.input.checked || automatic.input.checked;
		const integer = (input: HTMLInputElement) =>
			input.value !== '' &&
			Number.isInteger(Number(input.value)) &&
			Number(input.value) >= 1 &&
			Number(input.value) <= 32767;
		const valid =
			!enabled.input.checked ||
			(integer(start) &&
				integer(count) &&
				(automatic.input.checked ||
					(distance.value !== '' &&
						Number.isFinite(Number(distance.value)) &&
						Number(distance.value) >= 0 &&
						Number(distance.value) <= 22)));
		message.textContent = valid
			? ''
			: 'Enter whole numbers from 1 to 32767 and a distance from 0 to 22 inches.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [enabled.input, start, count, automatic.input, distance, restart]) {
		control.addEventListener('input', validate);
		control.addEventListener('change', validate);
	}
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (!host.canEdit() || !validate()) return;
		host.apply(
			enabled.input.checked
				? {
						start: Number(start.value),
						countBy: Number(count.value),
						restart: restart.value as LineNumberSettings['restart'],
						...(!automatic.input.checked
							? { distanceTwips: twips(Math.round(Number(distance.value) * 1440)) }
							: {}),
					}
				: undefined,
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
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return !element.hidden;
		},
		open() {
			const section = host.section();
			if (!section || !host.canEdit()) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const settings = section.lineNumberSettings;
			enabled.input.checked = Boolean(settings);
			start.value = String(settings?.start ?? 1);
			count.value = String(settings?.countBy ?? 1);
			automatic.input.checked = settings?.distanceTwips === undefined;
			distance.value = String((settings?.distanceTwips ?? 360) / 1440);
			restart.value = settings?.restart ?? 'continuous';
			validate();
			element.hidden = false;
			enabled.input.focus();
		},
	};
}
