import { dialogButton, fieldset, labelled, numberInput } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import type { ZoomFit } from './zoom-fit';

export interface ZoomHost {
	/** The current zoom, in percent. */
	percent(): number;
	setPercent(percent: number): void;
	fit(mode: ZoomFit): void;
	restoreFocus(): void;
}

type Choice = '200' | '100' | '75' | 'width' | 'page' | 'pages' | 'custom';
const CHOICES: ReadonlyArray<readonly [Choice, string]> = [
	['200', '200%'],
	['100', '100%'],
	['75', '75%'],
	['width', 'Page width'],
	['page', 'Whole page'],
	['pages', 'Two pages'],
	['custom', 'Percent:'],
];

/** The zoom the field holds, clamped to the 10 to 500 percent Word allows; null when it is not a number. */
export function parseZoomPercent(text: string): number | null {
	const value = Number(text.trim().replace(/%$/, ''));
	return text.trim() && Number.isFinite(value)
		? Math.max(10, Math.min(500, Math.round(value)))
		: null;
}

/**
 * Word's Zoom dialog: 200%, 100%, 75%, page width, whole page or a percentage, with a preview of
 * the text at that size. Text width, Wrap to window and page counts beyond two are not offered.
 */
export function createZoomDialog(host: ZoomHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-zoom-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Zoom settings');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Zoom';
	const radios = new Map<Choice, HTMLInputElement>();
	const options = CHOICES.map(([value, label]) => {
		const wrapper = document.createElement('label');
		wrapper.className = 'dve-dialog-check';
		const input = document.createElement('input');
		input.type = 'radio';
		input.name = 'dve-zoom';
		input.value = value;
		input.setAttribute('aria-label', label);
		const text = document.createElement('span');
		text.textContent = label;
		wrapper.append(input, text);
		radios.set(value, input);
		return wrapper;
	});
	const percent = numberInput(10, 500, 1);
	const preview = document.createElement('p');
	preview.className = 'dve-zoom-preview';
	preview.textContent = 'AaBbYyZz';
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, confirm);
	element.append(
		heading,
		fieldset('Zoom to', ...options, labelled('Percent', percent)),
		fieldset('Preview', preview),
		actions,
	);
	let locale: EditorLocale = 'en';
	const content = [...element.childNodes];
	element.replaceChildren();

	const chosen = (): Choice => [...radios].find(([, input]) => input.checked)?.[0] ?? 'custom';
	const showPreview = () => {
		const value = chosen();
		const size = value === 'custom' ? parseZoomPercent(percent.value) : Number(value);
		preview.style.fontSize = `${Math.round(((size ?? 100) / 100) * 14)}px`;
	};
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		const value = chosen();
		if (value === 'width' || value === 'page' || value === 'pages') host.fit(value);
		else {
			const next = value === 'custom' ? parseZoomPercent(percent.value) : Number(value);
			if (next === null) return percent.focus();
			host.setPercent(next);
		}
		hide();
	};
	for (const input of radios.values()) input.addEventListener('change', showPreview);
	percent.addEventListener('input', () => {
		const radio = radios.get('custom');
		if (radio) radio.checked = true;
		showPreview();
	});
	cancel.addEventListener('click', hide);
	confirm.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') hide();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		open() {
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const current = host.percent();
			const preset = (['200', '100', '75'] as const).find((value) => Number(value) === current);
			percent.value = String(current);
			const radio = radios.get(preset ?? 'custom');
			if (radio) radio.checked = true;
			showPreview();
			element.hidden = false;
			(radio ?? percent).focus();
		},
		close: hide,
		setLocale(next) {
			locale = next;
		},
		get isOpen() {
			return !element.hidden;
		},
	};
}
