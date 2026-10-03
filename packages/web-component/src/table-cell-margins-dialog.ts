import { signedTwips, type TableCellMargins } from 'docx-core';
import { checkbox, fieldset, labelled, numberInput, row, setTriState } from './dialog-fields';
import type { tablePropertiesContext } from './table-properties';

const sides = ['top', 'bottom', 'left', 'right'] as const;
type Side = (typeof sides)[number];

/** Selected-cell overrides stay separate from the table's defaults and untouched mixed values. */
export function createCellMarginFields(onChange: () => void) {
	const defaults = checkbox('Use table defaults');
	const controls = Object.fromEntries(
		sides.map((side) => [side, numberInput(0, 22, 0.005)]),
	) as Record<Side, HTMLInputElement>;
	const changed = new Set<Side>();
	let reset = false;
	const element = fieldset(
		'Selected cell margins (inches)',
		defaults.wrapper,
		row(labelled('Cell top', controls.top), labelled('Cell bottom', controls.bottom)),
		row(labelled('Cell left', controls.left), labelled('Cell right', controls.right)),
	);
	const syncDisabled = () => {
		for (const control of Object.values(controls)) control.disabled = defaults.input.checked;
	};
	defaults.input.addEventListener('change', () => {
		reset = defaults.input.checked;
		if (reset) changed.clear();
		syncDisabled();
		onChange();
	});
	for (const side of sides)
		controls[side].addEventListener('input', () => {
			changed.add(side);
			reset = false;
			defaults.input.checked = defaults.input.indeterminate = false;
			syncDisabled();
			onChange();
		});
	return {
		element,
		load(context: NonNullable<ReturnType<typeof tablePropertiesContext>>) {
			changed.clear();
			reset = false;
			const inherits = context.cells.map((cell) => !Object.keys(cell.margins).length);
			setTriState(
				defaults.input,
				inherits.every((value) => value === inherits[0]) ? inherits[0]! : null,
			);
			for (const side of sides) {
				const values = context.cells.map(
					(cell) =>
						cell.margins[side] ??
						context.margins[side] ??
						(side === 'left' || side === 'right' ? 108 : 0),
				);
				controls[side].value = values.every((value) => value === values[0])
					? String(values[0]! / 1440)
					: '';
			}
			syncDisabled();
		},
		valid() {
			return (
				reset ||
				[...changed].every((side) => {
					const value = Number(controls[side].value);
					return controls[side].value !== '' && Number.isFinite(value) && value >= 0 && value <= 22;
				})
			);
		},
		patch(): TableCellMargins | null | undefined {
			if (reset) return null;
			if (!changed.size) return undefined;
			return Object.fromEntries(
				[...changed].map((side) => [
					side,
					signedTwips(Math.round(Number(controls[side].value) * 1440)),
				]),
			);
		},
	};
}
