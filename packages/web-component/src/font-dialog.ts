import type { EditorView } from 'prosemirror-view';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	listInput,
	numberInput,
	row,
	selectOf,
	setTriState,
	textInput,
} from './dialog-fields';
import {
	applyFontFormat,
	readFontFormat,
	type FontFormat,
	type UnderlineKind,
} from './font-format';
import { focusView } from './focus-view';
import { localizeElement, type EditorLocale } from './localization';
import { FONT_FAMILIES, parseFontFamily, parseFontSize } from './ribbon-combo';

export interface FormatDialog {
	element: HTMLElement;
	open(): void;
	close(): void;
	setLocale(locale: EditorLocale): void;
	readonly isOpen: boolean;
}

type Style = 'regular' | 'italic' | 'bold' | 'boldItalic';
const UNDERLINES: ReadonlyArray<readonly [UnderlineKind, string]> = [
	['none', '(none)'],
	['single', 'Single'],
	['double', 'Double'],
	['dotted', 'Dotted'],
	['dash', 'Dashed'],
	['wave', 'Wave'],
];

const styleOf = (bold: boolean, italic: boolean): Style =>
	bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular';

/** Word's Font dialog: family, style, size, colour, underline, effects and character spacing. */
export function createFontDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-font-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Font');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Font';

	const { input: family, list } = listInput(FONT_FAMILIES);
	const style = selectOf([
		['regular', 'Regular'],
		['italic', 'Italic'],
		['bold', 'Bold'],
		['boldItalic', 'Bold Italic'],
	]);
	const size = textInput({ inputMode: 'decimal' });
	const color = Object.assign(document.createElement('input'), { type: 'color' });
	const underline = selectOf(UNDERLINES);
	const underlineColor = Object.assign(document.createElement('input'), { type: 'color' });
	const automatic = checkbox('Automatic');
	const effects = {
		strike: checkbox('Strikethrough'),
		doubleStrike: checkbox('Double strikethrough'),
		superscript: checkbox('Superscript'),
		subscript: checkbox('Subscript'),
		smallCaps: checkbox('Small caps'),
		caps: checkbox('All caps'),
		hidden: checkbox('Hidden'),
	};
	const spacingKind = selectOf([
		['normal', 'Normal'],
		['expanded', 'Expanded'],
		['condensed', 'Condensed'],
	]);
	const spacingBy = numberInput(0, 1584, 0.1);
	const preview = document.createElement('div');
	preview.className = 'dve-format-preview';
	preview.textContent = 'AaBbYyZz';
	preview.setAttribute('aria-hidden', 'true');

	const effectBoxes = Object.values(effects).map((item) => item.wrapper);
	element.append(
		heading,
		row(labelled('Font', family), labelled('Font style', style), labelled('Size', size)),
		row(
			labelled('Font color', color),
			labelled('Underline style', underline),
			labelled('Underline color', underlineColor),
		),
		automatic.wrapper,
		fieldset('Effects', ...effectBoxes),
		fieldset(
			'Character spacing',
			row(labelled('Character spacing', spacingKind), labelled('By', spacingBy)),
		),
		preview,
	);
	element.append(list);
	const cancel = dialogButton('Cancel');
	const confirm = dialogButton('OK', true);
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, confirm);
	element.append(actions);

	let initial: ReturnType<typeof readFontFormat> | undefined;
	const dirty = new Set<keyof FontFormat>();
	const mark = (...fields: Array<keyof FontFormat>) => {
		for (const field of fields) dirty.add(field);
		refreshPreview();
	};

	const refreshPreview = () => {
		const bold = style.value === 'bold' || style.value === 'boldItalic';
		const italic = style.value === 'italic' || style.value === 'boldItalic';
		const points = parseFontSize(size.value) ?? 11;
		const lines = [
			underline.value !== 'none' ? 'underline' : '',
			effects.strike.input.checked || effects.doubleStrike.input.checked ? 'line-through' : '',
		].filter(Boolean);
		const by = spacingKind.value === 'normal' ? 0 : Number(spacingBy.value) || 0;
		Object.assign(preview.style, {
			fontFamily: parseFontFamily(family.value) ? `"${family.value}", sans-serif` : '',
			fontSize: `${Math.min(30, Math.max(10, points * 1.3))}px`,
			fontWeight: bold ? '700' : '400',
			fontStyle: italic ? 'italic' : 'normal',
			color: color.value,
			textDecorationLine: lines.join(' '),
			textDecorationStyle:
				underline.value === 'double' || effects.doubleStrike.input.checked
					? 'double'
					: underline.value === 'dotted'
						? 'dotted'
						: underline.value === 'dash'
							? 'dashed'
							: underline.value === 'wave'
								? 'wavy'
								: 'solid',
			textDecorationColor: automatic.input.checked ? '' : underlineColor.value,
			textTransform: effects.caps.input.checked ? 'uppercase' : 'none',
			fontVariant: effects.smallCaps.input.checked ? 'small-caps' : 'normal',
			letterSpacing: spacingKind.value === 'condensed' ? `${-by}pt` : `${by}pt`,
			verticalAlign: effects.superscript.input.checked
				? 'super'
				: effects.subscript.input.checked
					? 'sub'
					: 'baseline',
			opacity: effects.hidden.input.checked ? '0.5' : '1',
		});
	};

	family.addEventListener('input', () => mark('family'));
	size.addEventListener('input', () => mark('size'));
	style.addEventListener('change', () => mark('bold', 'italic'));
	color.addEventListener('input', () => mark('color'));
	underline.addEventListener('change', () => mark('underline'));
	underlineColor.addEventListener('input', () => {
		automatic.input.checked = false;
		mark('underlineColor');
	});
	automatic.input.addEventListener('change', () => mark('underlineColor'));
	spacingKind.addEventListener('change', () => {
		if (spacingKind.value !== 'normal' && !Number(spacingBy.value)) spacingBy.value = '1';
		mark('spacing');
	});
	spacingBy.addEventListener('input', () => mark('spacing'));
	for (const [field, item] of Object.entries(effects))
		item.input.addEventListener('change', () => {
			item.input.indeterminate = false;
			if (item.input.checked && field === 'superscript') effects.subscript.input.checked = false;
			if (item.input.checked && field === 'subscript') effects.superscript.input.checked = false;
			mark(
				field === 'superscript' || field === 'subscript' ? 'script' : (field as keyof FontFormat),
			);
		});

	const collect = (): Partial<FontFormat> => {
		const changes: Record<string, unknown> = {};
		if (dirty.has('family')) {
			const name = parseFontFamily(family.value);
			if (name) changes.family = name;
		}
		if (dirty.has('size')) {
			const points = parseFontSize(size.value);
			if (points !== null) changes.size = points;
		}
		if (dirty.has('color')) changes.color = color.value;
		if (dirty.has('bold') || dirty.has('italic')) {
			changes.bold = style.value === 'bold' || style.value === 'boldItalic';
			changes.italic = style.value === 'italic' || style.value === 'boldItalic';
		}
		if (dirty.has('underline')) changes.underline = underline.value as UnderlineKind;
		if (dirty.has('underlineColor'))
			changes.underlineColor = automatic.input.checked ? null : underlineColor.value;
		for (const key of ['strike', 'doubleStrike', 'smallCaps', 'caps', 'hidden'] as const)
			if (dirty.has(key)) changes[key] = effects[key === 'hidden' ? 'hidden' : key].input.checked;
		if (dirty.has('script'))
			changes.script = effects.superscript.input.checked
				? 'superscript'
				: effects.subscript.input.checked
					? 'subscript'
					: 'none';
		if (dirty.has('spacing')) {
			const by = Number(spacingBy.value) || 0;
			changes.spacing =
				spacingKind.value === 'normal' ? 0 : spacingKind.value === 'condensed' ? -by : by;
		}
		return changes as Partial<FontFormat>;
	};

	const close = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	const submit = () => {
		const view = getView();
		if (view) applyFontFormat(view, collect());
		close();
	};
	cancel.addEventListener('click', close);
	confirm.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') close();
		else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});

	// The fields exist in the document only while the dialog is open, so their names never
	// collide with the ribbon's controls of the same name (Line spacing, Font size, ...).
	const content = [...element.childNodes];
	element.replaceChildren();
	let locale: EditorLocale = 'en';
	return {
		element,
		open() {
			const view = getView();
			if (!view?.editable) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			initial = readFontFormat(view.state);
			dirty.clear();
			family.value = initial.family ?? '';
			size.value = initial.size === null ? '' : String(initial.size);
			if (initial.bold === null || initial.italic === null) style.selectedIndex = -1;
			else style.value = styleOf(initial.bold, initial.italic);
			color.value = initial.color ?? '#000000';
			if (initial.underline === null) underline.selectedIndex = -1;
			else underline.value = initial.underline;
			automatic.input.checked = initial.underlineColor === null;
			underlineColor.value = initial.underlineColor ?? '#000000';
			setTriState(effects.strike.input, initial.strike);
			setTriState(effects.doubleStrike.input, initial.doubleStrike);
			setTriState(
				effects.superscript.input,
				initial.script === null ? null : initial.script === 'superscript',
			);
			setTriState(
				effects.subscript.input,
				initial.script === null ? null : initial.script === 'subscript',
			);
			setTriState(effects.smallCaps.input, initial.smallCaps);
			setTriState(effects.caps.input, initial.caps);
			setTriState(effects.hidden.input, initial.hidden);
			const spacing = initial.spacing ?? 0;
			spacingKind.value = spacing === 0 ? 'normal' : spacing > 0 ? 'expanded' : 'condensed';
			spacingBy.value = spacing ? String(Math.abs(spacing)) : '';
			refreshPreview();
			element.hidden = false;
			family.focus();
			family.select();
		},
		close,
		setLocale(next: EditorLocale) {
			locale = next;
		},
		get isOpen() {
			return !element.hidden;
		},
	};
}
