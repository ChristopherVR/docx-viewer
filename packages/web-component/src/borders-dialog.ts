import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, fieldset, labelled, row, selectOf } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	applyBordersAndShading,
	readBordersAndShading,
	type BordersAndShading,
} from './paragraph-decoration';

const WIDTHS: ReadonlyArray<readonly [string, string]> = [
	['2', '¼ pt'],
	['4', '½ pt'],
	['6', '¾ pt'],
	['8', '1 pt'],
	['12', '1½ pt'],
	['18', '2¼ pt'],
	['24', '3 pt'],
	['36', '4½ pt'],
	['48', '6 pt'],
];

/** Word's Borders and Shading dialog for paragraphs: sides, one pen (style, width, colour) and a fill. */
export function createBordersDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-borders-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Borders and Shading');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Borders and Shading';
	const sides = {
		top: checkbox('Top'),
		bottom: checkbox('Bottom'),
		left: checkbox('Left'),
		right: checkbox('Right'),
	};
	const style = selectOf([
		['single', 'Single'],
		['double', 'Double'],
		['dotted', 'Dotted'],
		['dashed', 'Dashed'],
	]);
	const width = selectOf(WIDTHS);
	const color = document.createElement('input');
	color.type = 'color';
	const automatic = checkbox('Automatic');
	const fill = document.createElement('input');
	fill.type = 'color';
	const noFill = checkbox('No Color');
	const ok = dialogButton('OK', true);
	const cancel = dialogButton('Cancel');
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	element.append(
		heading,
		fieldset(
			'Borders',
			row(...Object.values(sides).map((side) => side.wrapper)),
			row(labelled('Style', style), labelled('Width', width)),
			row(labelled('Color', color), automatic.wrapper),
		),
		fieldset('Shading', row(labelled('Fill', fill), noFill.wrapper)),
		actions,
	);
	let locale: EditorLocale = 'en';
	const content = [...element.childNodes];
	element.replaceChildren();

	const syncDisabled = () => {
		color.disabled = automatic.input.checked;
		fill.disabled = noFill.input.checked;
	};
	automatic.input.addEventListener('change', syncDisabled);
	noFill.input.addEventListener('change', syncDisabled);
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	ok.addEventListener('click', () => {
		const view = getView();
		if (!view) return;
		const settings: BordersAndShading = {
			sides: Object.fromEntries(
				Object.entries(sides).map(([name, box]) => [name, box.input.checked]),
			),
			style: style.value as BordersAndShading['style'],
			sizeEighthPoints: Number(width.value),
			color: automatic.input.checked ? null : color.value,
			fill: noFill.input.checked ? null : fill.value,
		};
		applyBordersAndShading(view, settings);
		hide();
	});
	cancel.addEventListener('click', hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') hide();
	});
	return {
		element,
		open() {
			const view = getView();
			if (!view) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			const current = readBordersAndShading(view.state);
			for (const [name, box] of Object.entries(sides))
				box.input.checked = Boolean(current.sides[name as 'top']);
			style.value = current.style;
			width.value = WIDTHS.some(([value]) => value === String(current.sizeEighthPoints))
				? String(current.sizeEighthPoints)
				: '4';
			automatic.input.checked = current.color === null;
			color.value = current.color ?? '#000000';
			noFill.input.checked = current.fill === null;
			fill.value = current.fill ?? '#ffff00';
			syncDisabled();
			ok.disabled = !view.editable;
			element.hidden = false;
			sides.top.input.focus();
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
