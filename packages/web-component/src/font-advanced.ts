import {
	checkbox,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	setTriState,
} from './dialog-fields';
import type { FontFormat, FontFormatState } from './font-format';
import { LIGATURE_VALUES, isLigatures } from '@christophervr/docx-core';
import { ligatureCss } from './ligature-style';

/** Word's Advanced character spacing controls. Untouched mixed values stay mixed. */
export function createFontAdvanced(mark: (...fields: Array<keyof FontFormat>) => void) {
	const scale = numberInput(0, 600, 1);
	const spacingKind = selectOf([
		['normal', 'Normal'],
		['expanded', 'Expanded'],
		['condensed', 'Condensed'],
	]);
	const spacingBy = numberInput(0, 1584, 0.05);
	const positionKind = selectOf([
		['normal', 'Normal'],
		['raised', 'Raised'],
		['lowered', 'Lowered'],
	]);
	const positionBy = numberInput(0, 1584, 0.5);
	const kerning = checkbox('Kerning for fonts');
	const kerningBy = numberInput(0.5, 1638, 0.5);
	const ligatures = selectOf(
		LIGATURE_VALUES.map((value) => [
			value,
			value === 'none'
				? 'None'
				: value === 'all'
					? 'All'
					: value.replace(/([A-Z])/g, ' $1').replace(/^./, (letter) => letter.toUpperCase()),
		]),
	);
	const element = fieldset(
		'Character spacing',
		row(labelled('Scale (%)', scale)),
		row(labelled('Character spacing', spacingKind), labelled('By', spacingBy)),
		row(labelled('Position', positionKind), labelled('Position by (points)', positionBy)),
		row(kerning.wrapper, labelled('Points and above', kerningBy)),
		row(labelled('Ligatures', ligatures)),
	);
	const note = document.createElement('p');
	note.textContent = 'Review advanced formatting in Print Layout.';
	element.append(note);
	for (const input of [scale, spacingBy, positionBy, kerningBy]) input.required = true;
	const sync = () => {
		spacingBy.disabled = spacingKind.value === 'normal' || spacingKind.selectedIndex < 0;
		positionBy.disabled = positionKind.value === 'normal' || positionKind.selectedIndex < 0;
		kerningBy.disabled = !kerning.input.checked || kerning.input.indeterminate;
	};
	scale.addEventListener('input', () => mark('scale'));
	ligatures.addEventListener('change', () => mark('ligatures'));
	spacingKind.addEventListener('change', () => {
		if (!Number(spacingBy.value)) spacingBy.value = '1';
		sync();
		mark('spacing');
	});
	spacingBy.addEventListener('input', () => mark('spacing'));
	positionKind.addEventListener('change', () => {
		if (!Number(positionBy.value)) positionBy.value = '3';
		sync();
		mark('position');
	});
	positionBy.addEventListener('input', () => mark('position'));
	kerning.input.addEventListener('change', () => {
		kerning.input.indeterminate = false;
		if (!Number(kerningBy.value)) kerningBy.value = '12';
		sync();
		mark('kerning');
	});
	kerningBy.addEventListener('input', () => mark('kerning'));
	return {
		element,
		spacingKind,
		spacingBy,
		reset(state: FontFormatState) {
			ligatures.value = state.ligatures ?? '';
			if (state.ligatures === null) ligatures.selectedIndex = -1;
			scale.value = state.scale === null ? '' : String(state.scale);
			const spacing = state.spacing ?? 0;
			spacingKind.value = spacing === 0 ? 'normal' : spacing > 0 ? 'expanded' : 'condensed';
			if (state.spacing === null) spacingKind.selectedIndex = -1;
			spacingBy.value = state.spacing === null ? '' : String(Math.abs(spacing));
			const position = state.position ?? 0;
			positionKind.value = position === 0 ? 'normal' : position > 0 ? 'raised' : 'lowered';
			if (state.position === null) positionKind.selectedIndex = -1;
			positionBy.value = state.position === null ? '' : String(Math.abs(position));
			setTriState(kerning.input, state.kerning === null ? null : state.kerning > 0);
			kerningBy.value = state.kerning === null ? '' : String(state.kerning || 12);
			sync();
		},
		valid(dirty: Set<keyof FontFormat>, reveal = () => {}) {
			for (const [key, input] of [
				['scale', scale],
				['spacing', spacingBy],
				['position', positionBy],
				['kerning', kerningBy],
			] as const)
				if (dirty.has(key) && !input.checkValidity()) {
					reveal();
					input.reportValidity();
					return false;
				}
			return true;
		},
		collect(dirty: Set<keyof FontFormat>): Partial<FontFormat> {
			const changes: Partial<FontFormat> = {};
			if (dirty.has('ligatures') && isLigatures(ligatures.value))
				changes.ligatures = ligatures.value;
			if (dirty.has('scale')) changes.scale = Number(scale.value);
			if (dirty.has('spacing'))
				changes.spacing =
					spacingKind.value === 'normal'
						? 0
						: Number(spacingBy.value) * (spacingKind.value === 'condensed' ? -1 : 1);
			if (dirty.has('position'))
				changes.position =
					positionKind.value === 'normal'
						? 0
						: Number(positionBy.value) * (positionKind.value === 'lowered' ? -1 : 1);
			if (dirty.has('kerning'))
				changes.kerning = kerning.input.checked ? Number(kerningBy.value) : 0;
			return changes;
		},
		previewStyle(size: number) {
			const factor = scale.value === '' ? 1 : Number(scale.value) / 100;
			const spacing =
				spacingKind.value === 'normal'
					? 0
					: Number(spacingBy.value) * (spacingKind.value === 'condensed' ? -1 : 1);
			return {
				fontVariantLigatures: isLigatures(ligatures.value)
					? ligatureCss(ligatures.value)
					: 'normal',
				transform: `scaleX(${factor})`,
				letterSpacing: `${factor ? spacing / factor : 0}pt`,
				position: 'relative',
				top: `${positionKind.value === 'normal' ? 0 : Number(positionBy.value) * (positionKind.value === 'raised' ? -1 : 1)}pt`,
				fontKerning: kerning.input.checked && size >= Number(kerningBy.value) ? 'normal' : 'none',
			};
		},
	};
}
