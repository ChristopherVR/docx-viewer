import type { EditorView } from 'prosemirror-view';
import { applyLink, bookmarkNames, linkAtSelection, removeLink } from './link-commands';
import { localizeElement, type EditorLocale } from './localization';
import { focusView } from './focus-view';

export interface LinkDialogOptions {
	getView: () => EditorView | undefined;
	onError: (error: Error) => void;
}

export interface LinkDialog {
	element: HTMLElement;
	open(): void;
	close(): void;
	setLocale(locale: EditorLocale): void;
	readonly isOpen: boolean;
}

function field(label: string, control: HTMLInputElement | HTMLSelectElement): HTMLLabelElement {
	const wrapper = document.createElement('label');
	const text = document.createElement('span');
	text.textContent = label;
	control.setAttribute('aria-label', label);
	wrapper.append(text, control);
	return wrapper;
}

function button(label: string, primary = false): HTMLButtonElement {
	const element = document.createElement('button');
	element.type = 'button';
	element.textContent = label;
	if (primary) element.className = 'dve-dialog-primary';
	return element;
}

/** Word's Insert/Edit Hyperlink dialog: web address or a bookmark in this document, plus a ScreenTip. */
export function createLinkDialog(options: LinkDialogOptions): LinkDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-link-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Insert link');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Insert link';
	const display = document.createElement('input');
	display.type = 'text';
	const address = document.createElement('input');
	address.type = 'url';
	address.placeholder = 'https://';
	const place = document.createElement('select');
	const tooltip = document.createElement('input');
	tooltip.type = 'text';
	const displayField = field('Text to display', display);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	const remove = button('Remove link');
	const cancel = button('Cancel');
	const confirm = button('Insert', true);
	actions.append(remove, cancel, confirm);
	element.append(
		heading,
		displayField,
		field('Address', address),
		field('Place in this document', place),
		field('ScreenTip', tooltip),
		actions,
	);

	const close = () => {
		element.hidden = true;
		focusView(options.getView());
	};
	const submit = () => {
		const view = options.getView();
		if (!view) return;
		const href = address.value.trim();
		const anchor = place.value;
		const tip = tooltip.value.trim();
		try {
			if (!href && !anchor) removeLink(view);
			else
				applyLink(
					view,
					{ ...(href ? { href } : { anchor }), ...(tip && { tooltip: tip }) },
					display.value.trim() || undefined,
				);
			close();
		} catch (cause) {
			options.onError(cause instanceof Error ? cause : new Error(String(cause)));
		}
	};
	address.addEventListener('input', () => {
		if (address.value) place.value = '';
	});
	place.addEventListener('change', () => {
		if (place.value) address.value = '';
	});
	remove.addEventListener('click', () => {
		const view = options.getView();
		if (view) removeLink(view);
		close();
	});
	cancel.addEventListener('click', close);
	confirm.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});

	return {
		element,
		open() {
			const view = options.getView();
			if (!view) return;
			const existing = linkAtSelection(view);
			const { from, to, empty } = view.state.selection;
			display.value = empty ? '' : view.state.doc.textBetween(from, to, ' ');
			displayField.hidden = !empty || Boolean(existing);
			address.value = existing?.href ?? '';
			tooltip.value = existing?.tooltip ?? '';
			place.replaceChildren(
				new Option('—', ''),
				...bookmarkNames(view).map((name) => new Option(name, name)),
			);
			place.value = existing?.anchor ?? '';
			remove.hidden = !existing;
			element.hidden = false;
			(existing?.anchor ? place : address).focus();
		},
		close,
		setLocale(locale) {
			localizeElement(element, locale);
		},
		get isOpen() {
			return !element.hidden;
		},
	};
}
