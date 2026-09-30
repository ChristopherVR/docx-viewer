import type { EditorView } from 'prosemirror-view';
import { dialogButton, fieldset, labelled, numberInput, row, selectOf } from './dialog-fields';
import { readDropCap, setDropCap, type DropCapStyle } from './drop-cap-command';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';

/** Word's position, font, height and text-distance controls for the selected initial. */
export function createDropCapDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Drop Cap Options');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Drop Cap Options';
	const position = selectOf([
		['none', 'None'],
		['drop', 'Dropped'],
		['margin', 'In margin'],
	]);
	const font = document.createElement('input');
	font.type = 'text';
	font.maxLength = 255;
	const lines = numberInput(1, 10, 1);
	const distance = numberInput(0, 22, 0.01);
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const content = [
		heading,
		fieldset('Position', labelled('Position', position)),
		fieldset(
			'Options',
			labelled('Font family', font),
			row(labelled('Lines to drop', lines), labelled('Distance from text (inches)', distance)),
		),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	const validate = () => {
		const disabled = position.value === 'none';
		for (const control of [font, lines, distance]) control.disabled = disabled;
		const valid =
			disabled ||
			(font.value.trim().length > 0 &&
				lines.value !== '' &&
				Number.isInteger(Number(lines.value)) &&
				Number(lines.value) >= 1 &&
				Number(lines.value) <= 10 &&
				distance.value !== '' &&
				Number.isFinite(Number(distance.value)) &&
				Number(distance.value) >= 0 &&
				Number(distance.value) <= 22);
		message.textContent = valid
			? ''
			: 'Enter a font, 1 to 10 lines and a distance from 0 to 22 inches.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [position, font, lines, distance]) {
		control.addEventListener('input', validate);
		control.addEventListener('change', validate);
	}
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		if (!view?.editable || !validate()) return;
		if (
			setDropCap(view, position.value as DropCapStyle, {
				lines: Number(lines.value),
				distanceTwips: Math.round(Number(distance.value) * 1440),
				fontFamily: font.value.trim(),
			}) ||
			position.value === 'none'
		)
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
			const view = getView();
			if (!view?.editable) return;
			const settings = readDropCap(view);
			if (!settings) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			position.value = settings.style;
			font.value = settings.fontFamily;
			lines.value = String(settings.lines);
			distance.value = String(settings.distanceTwips / 1440);
			validate();
			element.hidden = false;
			position.focus();
		},
	};
}
