import type { SectionProperties } from '@christophervr/docx-core';
import { dialogButton, fieldset, labelled, numberInput, row, selectOf } from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	readPageSetup,
	validatePageSetup,
	presetOf,
	type PageSetupProblem,
	type PageSetupValues,
} from './page-setup-model';
import { PAGE_SIZES, PAGE_SIZE_OPTIONS } from './page-size';

export interface PageSetupHost {
	/** The section holding the selection, or undefined when there is none to edit. */
	section(): SectionProperties | undefined;
	canEdit(): boolean;
	apply(values: PageSetupValues): void;
	restoreFocus(): void;
}

const MESSAGES: Record<PageSetupProblem, string> = {
	range: 'Enter measurements from 0 to 22 inches (a page must be larger than 0).',
	width:
		'The margins leave less than 0.5 inch of text width. Reduce the left, right or gutter margin.',
	height: 'The margins leave less than 0.5 inch of text height. Reduce the top or bottom margin.',
};

/** Word's Page Setup dialog: margins, gutter, header and footer distances, orientation and paper. */
export function createPageSetupDialog(host: PageSetupHost): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-page-setup-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Page setup');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Page setup';

	const inch = () => numberInput(0, 22, 0.05);
	const top = inch();
	const bottom = inch();
	const left = inch();
	const right = inch();
	const gutter = inch();
	const header = inch();
	const footer = inch();
	const orientation = selectOf([
		['portrait', 'Portrait'],
		['landscape', 'Landscape'],
	]);
	const size = selectOf([...PAGE_SIZE_OPTIONS, ['custom', 'Custom size']]);
	const width = numberInput(0.1, 22, 0.05);
	const height = numberInput(0.1, 22, 0.05);
	const message = document.createElement('p');
	message.className = 'dve-dialog-message';
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, confirm);
	element.append(
		heading,
		fieldset(
			'Margins',
			row(labelled('Top', top), labelled('Bottom', bottom)),
			row(labelled('Left', left), labelled('Right', right)),
			row(labelled('Gutter', gutter)),
		),
		fieldset(
			'Header',
			row(labelled('Header from edge', header), labelled('Footer from edge', footer)),
		),
		fieldset(
			'Paper size',
			row(labelled('Orientation', orientation), labelled('Paper size', size)),
			row(labelled('Width', width), labelled('Height', height)),
		),
		message,
		actions,
	);
	let locale: EditorLocale = 'en';
	const content = [...element.childNodes];
	element.replaceChildren();

	const values = (): PageSetupValues => ({
		topIn: Number(top.value),
		bottomIn: Number(bottom.value),
		leftIn: Number(left.value),
		rightIn: Number(right.value),
		gutterIn: Number(gutter.value),
		headerIn: Number(header.value),
		footerIn: Number(footer.value),
		orientation: orientation.value === 'landscape' ? 'landscape' : 'portrait',
		widthIn: Number(width.value),
		heightIn: Number(height.value),
	});
	const say = (problem: PageSetupProblem | null) => {
		message.textContent = problem ? MESSAGES[problem] : '';
		localizeElement(message, locale);
	};
	const validate = () => {
		const problem = validatePageSetup(values());
		say(problem);
		confirm.disabled = Boolean(problem);
	};
	size.addEventListener('change', () => {
		const preset = PAGE_SIZES[size.value];
		if (preset) {
			width.value = String(Math.round((preset[0] / 1440) * 1000) / 1000);
			height.value = String(Math.round((preset[1] / 1440) * 1000) / 1000);
		}
		validate();
	});
	const syncPreset = () => {
		size.value = presetOf(values()) ?? 'custom';
		validate();
	};
	width.addEventListener('input', syncPreset);
	height.addEventListener('input', syncPreset);
	for (const input of [top, bottom, left, right, gutter, header, footer])
		input.addEventListener('input', validate);
	orientation.addEventListener('change', validate);

	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		host.restoreFocus();
	};
	const submit = () => {
		if (validatePageSetup(values())) return;
		host.apply(values());
		hide();
	};
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
			const section = host.section();
			if (!section || !host.canEdit()) return;
			const current = readPageSetup(section);
			element.replaceChildren(...content);
			localizeElement(element, locale);
			top.value = String(current.topIn);
			bottom.value = String(current.bottomIn);
			left.value = String(current.leftIn);
			right.value = String(current.rightIn);
			gutter.value = String(current.gutterIn);
			header.value = String(current.headerIn);
			footer.value = String(current.footerIn);
			orientation.value = current.orientation;
			width.value = String(current.widthIn);
			height.value = String(current.heightIn);
			syncPreset();
			element.hidden = false;
			top.focus();
			top.select();
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
