import type { EditorView } from 'prosemirror-view';
import { focusView } from './focus-view';
import { localizeElement, type EditorLocale } from './localization';
import { schema } from './schema';

export interface PictureDialog {
	element: HTMLElement;
	open(pos: number): void;
	close(): void;
	setLocale(locale: EditorLocale): void;
	readonly isOpen: boolean;
}

function labelled(label: string, control: HTMLElement): HTMLLabelElement {
	const wrapper = document.createElement('label');
	const text = document.createElement('span');
	text.textContent = label;
	control.setAttribute('aria-label', label);
	wrapper.append(text, control);
	return wrapper;
}

function sizeInput(): HTMLInputElement {
	const input = document.createElement('input');
	input.type = 'number';
	input.min = '16';
	input.step = '1';
	return input;
}

/** Word's Format Picture essentials: alt text and size (px) with an aspect-ratio lock. */
export function createPictureDialog(getView: () => EditorView | undefined): PictureDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-picture-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Format picture');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Format picture';
	const alt = document.createElement('textarea');
	alt.rows = 3;
	const width = sizeInput();
	const height = sizeInput();
	const lock = document.createElement('input');
	lock.type = 'checkbox';
	lock.checked = true;
	const lockLabel = document.createElement('label');
	lockLabel.className = 'dve-dialog-check';
	const lockText = document.createElement('span');
	lockText.textContent = 'Lock aspect ratio';
	lockLabel.append(lock, lockText);
	const sizes = document.createElement('div');
	sizes.className = 'dve-dialog-row';
	sizes.append(labelled('Width', width), labelled('Height', height));
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	const cancel = document.createElement('button');
	cancel.type = 'button';
	cancel.textContent = 'Cancel';
	const apply = document.createElement('button');
	apply.type = 'button';
	apply.className = 'dve-dialog-primary';
	apply.textContent = 'OK';
	actions.append(cancel, apply);
	element.append(heading, labelled('Alt text', alt), sizes, lockLabel, actions);

	let position = -1;
	let ratio = 1;
	width.addEventListener('input', () => {
		if (lock.checked && Number(width.value) > 0)
			height.value = String(Math.round(Number(width.value) * ratio));
	});
	height.addEventListener('input', () => {
		if (lock.checked && Number(height.value) > 0)
			width.value = String(Math.round(Number(height.value) / ratio));
	});
	const close = () => {
		element.hidden = true;
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		const node = view?.state.doc.nodeAt(position);
		if (!view || !node || node.type !== schema.nodes.image) return close();
		const nextWidth = Math.max(16, Math.round(Number(width.value) || node.attrs.widthPx));
		const nextHeight = Math.max(16, Math.round(Number(height.value) || node.attrs.heightPx));
		view.dispatch(
			view.state.tr.setNodeMarkup(position, undefined, {
				...node.attrs,
				altText: alt.value.trim() || null,
				widthPx: nextWidth,
				heightPx: nextHeight,
			}),
		);
		close();
	};
	cancel.addEventListener('click', close);
	apply.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		open(pos) {
			const node = getView()?.state.doc.nodeAt(pos);
			if (!node || node.type !== schema.nodes.image) return;
			position = pos;
			ratio = Number(node.attrs.heightPx) / Math.max(1, Number(node.attrs.widthPx));
			alt.value = node.attrs.altText ?? '';
			width.value = String(node.attrs.widthPx);
			height.value = String(node.attrs.heightPx);
			element.hidden = false;
			alt.focus();
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
