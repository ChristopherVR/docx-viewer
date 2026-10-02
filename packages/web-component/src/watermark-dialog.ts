import type { WatermarkSpec } from 'docx-core';
import {
	checkbox,
	dialogButton,
	labelled,
	listInput,
	row,
	selectOf,
	textInput,
} from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';

export interface WatermarkHost {
	current(): WatermarkSpec | undefined;
	canEdit(): boolean;
	/** Sets the watermark in every header; undefined removes it. */
	apply(spec: WatermarkSpec | undefined): void;
	restoreFocus(): void;
}

const PRESETS = ['CONFIDENTIAL', 'DRAFT', 'DO NOT COPY', 'SAMPLE', 'URGENT'];
const DEFAULT_COLOR = '#c0c0c0';

/** Word's Printed Watermark for text: the words, font, colour, layout and transparency. */
export function createWatermarkDialog(host: WatermarkHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-watermark-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Printed watermark');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Printed watermark';
	const setting = selectOf([
		['none', 'No watermark'],
		['text', 'Text watermark'],
	]);
	const { input: text, list } = listInput(PRESETS);
	const font = textInput();
	const color = document.createElement('input');
	color.type = 'color';
	const layout = selectOf([
		['diagonal', 'Diagonal'],
		['horizontal', 'Horizontal'],
	]);
	const semi = checkbox('Semitransparent');
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
		labelled('Text', text),
		list,
		row(labelled('Font', font), labelled('Color', color)),
		row(labelled('Layout', layout), semi.wrapper),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	const validate = () => {
		const on = setting.value === 'text';
		for (const control of [text, font, color, layout, semi.input]) control.disabled = !on;
		const valid = !on || text.value.trim() !== '';
		message.textContent = valid ? '' : 'Enter the watermark text.';
		localizeElement(message, locale);
		ok.disabled = !valid;
		return valid;
	};
	for (const control of [setting, text, font, color, layout, semi.input])
		for (const type of ['input', 'change']) control.addEventListener(type, validate);
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (!host.canEdit() || !validate()) return;
		host.apply(
			setting.value === 'text'
				? {
						text: text.value.trim(),
						color: color.value.toLowerCase(),
						semitransparent: semi.input.checked,
						layout: layout.value as WatermarkSpec['layout'],
						...(font.value.trim() ? { fontFamily: font.value.trim() } : {}),
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
			if (!host.canEdit()) return;
			const current = host.current();
			element.replaceChildren(...content);
			localizeElement(element, locale);
			setting.value = current ? 'text' : 'none';
			text.value = current?.text ?? '';
			font.value = current?.fontFamily ?? 'Calibri';
			color.value = current?.color ?? DEFAULT_COLOR;
			layout.value = current?.layout ?? 'diagonal';
			semi.input.checked = current?.semitransparent ?? true;
			validate();
			element.hidden = false;
			setting.focus();
		},
	};
}
