import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, labelled, numberInput, row } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	insertTextBox,
	selectedTextBox,
	updateTextBox,
	type TextBoxSettings,
} from './text-box-commands';

const DEFAULT: TextBoxSettings = { lines: [], widthPx: 192, heightPx: 96, border: true };
const inches = (px: number) => String(Math.round((px / 96) * 100) / 100);

/** Insert > Text Box: inserts an inline box at the selection, or edits the selected simple box. */
export function createTextBoxDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-text-box-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Text box');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Text box';
	const text = document.createElement('textarea');
	text.rows = 5;
	const width = numberInput(0.25, 22, 0.01);
	const height = numberInput(0.25, 22, 0.01);
	const border = checkbox('Outline');
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const content = [
		heading,
		labelled('Text', text),
		row(labelled('Width (inches)', width), labelled('Height (inches)', height), border.wrapper),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	let editing: number | undefined;
	const inRange = (input: HTMLInputElement) =>
		input.value !== '' && Number(input.value) >= 0.25 && Number(input.value) <= 22;
	const validate = () => {
		const valid = inRange(width) && inRange(height);
		message.textContent = valid ? '' : 'Enter a width and height from 0.25 to 22 inches.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [width, height]) control.addEventListener('input', validate);
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		if (!view?.editable || !validate()) return;
		const settings: TextBoxSettings = {
			lines: text.value.split(/\r\n|\r|\n/),
			widthPx: Math.round(Number(width.value) * 96),
			heightPx: Math.round(Number(height.value) * 96),
			border: border.input.checked,
		};
		if (editing !== undefined) updateTextBox(view, editing, settings);
		else insertTextBox(view, settings);
		hide();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			hide();
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
			const view = getView();
			if (!view?.editable) return;
			const selected = selectedTextBox(view.state);
			editing = selected?.pos;
			const settings = selected?.settings ?? DEFAULT;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			text.value = settings.lines.join('\n');
			width.value = inches(settings.widthPx);
			height.value = inches(settings.heightPx);
			border.input.checked = settings.border;
			validate();
			element.hidden = false;
			text.focus();
		},
	};
}
