import type { TableCell } from 'docx-core';
import { fieldset, labelled, selectOf } from './dialog-fields';
import type { tablePropertiesContext } from './table-properties';

/** Exposes editable alignments without normalizing untouched imported or mixed values. */
export function createCellAlignmentField() {
	const control = selectOf([
		['', 'Mixed'],
		['default', 'Default alignment'],
		['top', 'Top'],
		['center', 'Center'],
		['bottom', 'Bottom'],
		['both', 'Justified (imported)'],
	]);
	const mixed = control.options[0]!;
	mixed.disabled = true;
	const justified = control.options[control.options.length - 1]!;
	justified.disabled = true;
	let changed = false;
	control.addEventListener('change', () => {
		changed = true;
	});
	return {
		element: fieldset('Cell alignment', labelled('Cell vertical alignment', control)),
		load(context: NonNullable<ReturnType<typeof tablePropertiesContext>>) {
			changed = false;
			const values = context.cells.map((cell) => cell.verticalAlign ?? 'default');
			control.value = values.every((value) => value === values[0]) ? values[0]! : '';
			mixed.hidden = control.value !== '';
			justified.hidden = control.value !== 'both';
		},
		patch(): TableCell['verticalAlign'] | null | undefined {
			if (!changed) return undefined;
			return control.value === 'default'
				? null
				: (control.value as NonNullable<TableCell['verticalAlign']>);
		},
	};
}
