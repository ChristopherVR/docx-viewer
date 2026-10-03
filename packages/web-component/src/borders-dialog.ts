import type { EditorView } from 'prosemirror-view';
import { checkbox, dialogButton, fieldset, labelled, row, selectOf } from './dialog-fields';
import { readCellFill } from './table-shading';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';
import {
	applyCellBorderSettings,
	readCellBorderSettings,
	tableBorderContext,
	type CellBorderSettings,
} from './table-border-commands';
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
	const inside = { insideH: checkbox('Inside horizontal'), insideV: checkbox('Inside vertical') };
	const scope = selectOf([
		['paragraph', 'Paragraph'],
		['cell', 'Cell'],
		['table', 'Table'],
	]);
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
			labelled('Apply to', scope),
			row(...Object.values(sides).map((side) => side.wrapper)),
			row(inside.insideH.wrapper, inside.insideV.wrapper),
			row(labelled('Style', style), labelled('Width', width)),
			row(labelled('Color', color), automatic.wrapper),
		),
		fieldset('Shading', row(labelled('Fill', fill), noFill.wrapper)),
		actions,
	);
	let locale: EditorLocale = 'en';
	let fillChanged = false;
	const content = [...element.childNodes];
	element.replaceChildren();

	const loadParagraph = (view: EditorView) => {
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
	};
	const inTable = () => {
		const view = getView();
		return Boolean(view && tableBorderContext(view.state));
	};
	const show = () => {
		const view = getView();
		if (!view) return;
		const cells = scope.value !== 'paragraph';
		inside.insideH.wrapper.hidden = !cells;
		inside.insideV.wrapper.hidden = !cells;
		if (!cells) loadParagraph(view);
		fillChanged = false;
		noFill.input.indeterminate = false;
		if (cells) {
			const current = readCellFill(
				tableBorderContext(view.state)!,
				view.state,
				scope.value as 'cell' | 'table',
			);
			noFill.input.checked = current === null;
			noFill.input.indeterminate = current === undefined;
			fill.value = current ?? '#ffff00';
			const read = readCellBorderSettings(view.state, scope.value as 'cell' | 'table');
			if (!read) return;
			for (const [name, box] of Object.entries(sides))
				box.input.checked = Boolean(read.sides[name as 'top']);
			inside.insideH.input.checked = Boolean(read.sides.insideH);
			inside.insideV.input.checked = Boolean(read.sides.insideV);
			style.value = ['single', 'double', 'dotted', 'dashed'].includes(read.pen.style)
				? read.pen.style
				: 'single';
			width.value = WIDTHS.some(([value]) => value === String(read.pen.sizeEighthPoints))
				? String(read.pen.sizeEighthPoints)
				: '4';
			automatic.input.checked = !read.pen.color;
			color.value = read.pen.color ?? '#000000';
		}
		syncDisabled();
	};
	scope.addEventListener('change', show);
	const syncDisabled = () => {
		color.disabled = automatic.input.checked;
		fill.disabled = noFill.input.checked;
	};
	automatic.input.addEventListener('change', syncDisabled);
	noFill.input.addEventListener('change', () => {
		fillChanged = true;
		syncDisabled();
	});
	fill.addEventListener('input', () => {
		fillChanged = true;
		noFill.input.checked = false;
		noFill.input.indeterminate = false;
		syncDisabled();
	});
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	ok.addEventListener('click', () => {
		const view = getView();
		if (!view) return;
		if (scope.value !== 'paragraph') {
			const cellSettings: CellBorderSettings = {
				scope: scope.value as 'cell' | 'table',
				...(fillChanged ? { fill: noFill.input.checked ? null : fill.value } : {}),
				sides: {
					...Object.fromEntries(
						Object.entries(sides).map(([name, box]) => [name, box.input.checked]),
					),
					insideH: inside.insideH.input.checked,
					insideV: inside.insideV.input.checked,
				},
				pen: {
					style: style.value,
					sizeEighthPoints: Number(width.value),
					...(automatic.input.checked ? {} : { color: color.value }),
				},
			};
			applyCellBorderSettings(view, cellSettings);
			hide();
			return;
		}
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
			const tabled = inTable();
			scope.parentElement!.hidden = !tabled;
			scope.value = tabled ? 'cell' : 'paragraph';
			loadParagraph(view);
			show();
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
