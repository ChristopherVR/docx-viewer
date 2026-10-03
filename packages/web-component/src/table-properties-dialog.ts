import { signedTwips, twips, type TableCellMargins } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	setTriState,
} from './dialog-fields';
import type { FormatDialog } from './font-dialog';
import { focusView } from './focus-view';
import { createCellMarginFields } from './table-cell-margins-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	applyTableProperties,
	tablePropertiesContext,
	type TableRowProperties,
	type RowPatch,
} from './table-properties';

export function createTablePropertiesDialog(viewOf: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Table properties');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Table properties';
	const specified = checkbox('Specify height');
	const height = numberInput(0.01, 22, 0.01);
	const rule = selectOf([
		['atLeast', 'At least'],
		['exact', 'Exactly'],
	]);
	const split = checkbox('Allow row to break across pages');
	const header = checkbox('Repeat as header row at the top of each page');
	const sides = ['top', 'bottom', 'left', 'right'] as const;
	const margins = Object.fromEntries(
		sides.map((side) => [side, numberInput(0, 22, 0.005)]),
	) as Record<(typeof sides)[number], HTMLInputElement>;
	const note = document.createElement('p');
	note.textContent =
		'Row settings apply to selected rows. Cell margins apply to selected cells; untouched overrides are kept.';
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const cancel = dialogButton('Cancel');
	const ok = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const cellMargins = createCellMarginFields(() => validate());
	const content = [
		heading,
		fieldset(
			'Row',
			specified.wrapper,
			row(labelled('Height (inches)', height), labelled('Row height is', rule)),
			split.wrapper,
			header.wrapper,
		),
		fieldset(
			'Default cell margins (inches)',
			row(labelled('Top', margins.top), labelled('Bottom', margins.bottom)),
			row(labelled('Left', margins.left), labelled('Right', margins.right)),
		),
		cellMargins.element,
		note,
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	const changed = new Set<string>();
	const controls = {
		specified: specified.input,
		height,
		rule,
		split: split.input,
		header: header.input,
		...margins,
	};
	const validate = () => {
		height.disabled = rule.disabled = !specified.input.checked;
		const invalid = Object.entries(controls).some(
			([key, control]) =>
				changed.has(key) &&
				control instanceof HTMLInputElement &&
				control.type === 'number' &&
				(key !== 'height' || specified.input.checked) &&
				(control.value === '' ||
					!Number.isFinite(Number(control.value)) ||
					Number(control.value) < Number(control.min) ||
					Number(control.value) > Number(control.max)),
		);
		const invalidCell = !cellMargins.valid();
		ok.disabled = invalid || invalidCell;
		message.textContent =
			invalid || invalidCell ? 'Enter a measurement within the displayed range.' : '';
		localizeElement(message, locale);
		return !invalid && !invalidCell;
	};
	for (const [key, control] of Object.entries(controls)) {
		control.addEventListener('input', () => {
			changed.add(key);
			validate();
		});
		control.addEventListener('change', () => {
			changed.add(key);
			validate();
		});
	}
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(viewOf());
	};
	const submit = () => {
		const view = viewOf();
		if (!view || !validate()) return;
		const patch: RowPatch = {};
		if (changed.has('specified') || changed.has('height') || changed.has('rule')) {
			if (
				specified.input.checked &&
				(changed.has('specified') || changed.has('height')) &&
				(!height.value || Number(height.value) <= 0)
			) {
				changed.add('height');
				validate();
				return;
			}
			if (!specified.input.checked) {
				patch.heightTwips = undefined;
				patch.heightRule = undefined;
			} else {
				if (changed.has('specified') || changed.has('height'))
					patch.heightTwips = twips(Math.round(Number(height.value) * 1440));
				if (rule.value && (changed.has('rule') || changed.has('specified')))
					patch.heightRule = rule.value === 'exact' ? 'exact' : 'atLeast';
			}
		}
		if (changed.has('split')) patch.cantSplit = !split.input.checked;
		if (changed.has('header')) patch.header = header.input.checked;
		const marginPatch: TableCellMargins = {};
		for (const side of sides)
			if (changed.has(side))
				marginPatch[side] = signedTwips(Math.round(Number(margins[side].value) * 1440));
		if (applyTableProperties(view, patch, marginPatch, cellMargins.patch())) hide();
	};
	cancel.addEventListener('click', hide);
	ok.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			hide();
		}
		if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
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
			const view = viewOf();
			const context = view && tablePropertiesContext(view.state);
			if (!view?.editable || !context) return;
			changed.clear();
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const common = <T>(read: (properties: TableRowProperties) => T): T | null => {
				const values = context.rows.map((row) => read(row.properties));
				return values.every((value) => value === values[0]) ? values[0]! : null;
			};
			setTriState(
				specified.input,
				common((p) => p.heightTwips != null),
			);
			const heightTwips = common((p) => p.heightTwips ?? null);
			height.value = heightTwips == null ? '' : String(heightTwips / 1440);
			rule.value = common((p) => p.heightRule ?? 'atLeast') ?? '';
			setTriState(
				split.input,
				common((p) => !p.cantSplit),
			);
			setTriState(
				header.input,
				common((p) => Boolean(p.header)),
			);
			for (const side of sides)
				margins[side].value = String(
					(context.margins[side] ?? (side === 'left' || side === 'right' ? 108 : 0)) / 1440,
				);
			cellMargins.load(context);
			validate();
			element.hidden = false;
			specified.input.focus();
		},
	};
}
