/** Small builders for the ribbon's format dialogs (Font, Paragraph). Text stays English until localized. */

export function labelled<T extends HTMLElement>(label: string, control: T): HTMLLabelElement {
	const wrapper = document.createElement('label');
	const text = document.createElement('span');
	text.textContent = label;
	control.setAttribute('aria-label', label);
	wrapper.append(text, control);
	return wrapper;
}

export function textInput(options: { inputMode?: string } = {}) {
	const input = document.createElement('input');
	input.type = 'text';
	if (options.inputMode) input.inputMode = options.inputMode;
	return input;
}

/** A text input offering `values` through a native list (Word's editable drop-down). */
export function listInput(values: readonly string[]): {
	input: HTMLInputElement;
	list: HTMLDataListElement;
} {
	const input = textInput();
	const list = document.createElement('datalist');
	list.id = `dve-list-${Math.random().toString(36).slice(2, 8)}`;
	for (const value of values) list.append(new Option(value));
	input.setAttribute('list', list.id);
	return { input, list };
}

export function numberInput(min: number, max: number, step: number) {
	const input = document.createElement('input');
	input.type = 'number';
	input.min = String(min);
	input.max = String(max);
	input.step = String(step);
	return input;
}

export function selectOf(options: ReadonlyArray<readonly [value: string, label: string]>) {
	const select = document.createElement('select');
	for (const [value, label] of options) select.append(new Option(label, value));
	return select;
}

/** A checkbox row whose text is its accessible name. */
export function checkbox(label: string) {
	const wrapper = document.createElement('label');
	wrapper.className = 'dve-dialog-check';
	const input = document.createElement('input');
	input.type = 'checkbox';
	const text = document.createElement('span');
	text.textContent = label;
	input.setAttribute('aria-label', label);
	wrapper.append(input, text);
	return { wrapper, input };
}

export function fieldset(legend: string, ...children: HTMLElement[]) {
	const set = document.createElement('fieldset');
	set.className = 'dve-dialog-fieldset';
	const title = document.createElement('legend');
	title.textContent = legend;
	set.append(title, ...children);
	return set;
}

export function row(...children: HTMLElement[]) {
	const el = document.createElement('div');
	el.className = 'dve-dialog-row dve-dialog-cols';
	el.style.setProperty('--cols', String(children.length));
	el.append(...children);
	return el;
}

export function dialogButton(label: string, primary = false) {
	const button = document.createElement('button');
	button.type = 'button';
	button.textContent = label;
	if (primary) button.className = 'dve-dialog-primary';
	return button;
}

/** A tri-state checkbox: `null` shows the indeterminate mark (the selection mixes values). */
export function setTriState(input: HTMLInputElement, value: boolean | null): void {
	input.indeterminate = value === null;
	input.checked = value === true;
}
