import {
	ensureListDefinition,
	formatListNumber,
	resolveNumberingLevel,
	signedTwips,
	twips,
	type DocumentModel,
	type NumberingLevelDefinition,
} from '@christophervr/docx-core';
import type { EditorView } from 'prosemirror-view';
import { applyCustomList } from './list-commands';
import { nodeNumbering } from './node-numbering';
import {
	checkbox,
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	textInput,
} from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, translate, type EditorLocale } from './localization';

/** Creates a fresh definition for the selection; imported definitions remain untouched. */
export function createMultilevelListDialog(
	getView: () => EditorView | undefined,
	getModel: () => DocumentModel,
	canEdit = () => true,
): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Define New Multilevel List');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Define New Multilevel List';
	const level = selectOf(Array.from({ length: 9 }, (_, i) => [String(i), String(i + 1)] as const));
	const style = selectOf([
		['decimal', '1, 2, 3, …'],
		['decimalZero', '01, 02, 03, …'],
		['upperRoman', 'I, II, III, …'],
		['lowerRoman', 'i, ii, iii, …'],
		['upperLetter', 'A, B, C, …'],
		['lowerLetter', 'a, b, c, …'],
		['ordinal', '1st, 2nd, 3rd, …'],
		['cardinalText', 'One, Two, Three, …'],
		['ordinalText', 'First, Second, Third, …'],
		['bullet', 'Bullet'],
		['none', 'None'],
	]);
	const pattern = textInput();
	pattern.maxLength = 255;
	const start = numberInput(0, 32767, 1);
	const restart = selectOf([]);
	const legal = checkbox('Legal style numbering');
	const alignment = selectOf([
		['left', 'Left'],
		['center', 'Center'],
		['right', 'Right'],
	]);
	const aligned = numberInput(-22, 22, 0.01);
	const indent = numberInput(-22, 22, 0.01);
	const suffix = selectOf([
		['tab', 'Tab character'],
		['space', 'Space'],
		['none', 'Nothing'],
	]);
	const linkedStyle = selectOf([['', '(No style)']]);
	const tabStop = numberInput(0, 22, 0.01);
	const markerFont = textInput();
	const markerSize = numberInput(1, 400, 0.5);
	const markerColor = textInput();
	markerColor.placeholder = '#rrggbb';
	const markerBold = checkbox('Bold');
	const markerItalic = checkbox('Italic');
	const hint = document.createElement('p');
	hint.style.cssText = 'font-size:12px;margin:4px 0';
	hint.textContent = 'Use %1 through %9 for level numbers. Creates a new list for the selection.';
	const preview = document.createElement('div');
	preview.setAttribute('role', 'group');
	preview.setAttribute('aria-label', 'List preview');
	preview.style.cssText =
		'max-height:70px;overflow:auto;padding:8px;border:1px solid var(--dve-border,#ccc)';
	const message = document.createElement('p');
	message.setAttribute('role', 'alert');
	const ok = dialogButton('OK', true),
		cancel = dialogButton('Cancel');
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(cancel, ok);
	const levelField = labelled('Level to modify', level);
	levelField.style.alignContent = 'start';
	const content = [
		heading,
		row(levelField, preview),
		fieldset(
			'Number format',
			row(
				labelled('Number style for this level', style),
				labelled('Start at', start),
				labelled('Link level to style', linkedStyle),
			),
			labelled('Enter formatting for number', pattern),
			hint,
			row(
				labelled('Restart list after', restart),
				legal.wrapper,
				markerBold.wrapper,
				markerItalic.wrapper,
			),
		),
		fieldset(
			'Position',
			row(
				labelled('Number alignment', alignment),
				labelled('Aligned at (inches)', aligned),
				labelled('Text indent at (inches)', indent),
				labelled('Add tab stop at (inches)', tabStop),
			),
			row(
				labelled('Follow number with', suffix),
				labelled('Font', markerFont),
				labelled('Size (pt)', markerSize),
				labelled('Color (#rrggbb, blank for automatic)', markerColor),
			),
		),
		message,
		actions,
	];
	let locale: EditorLocale = 'en';
	let drafts: NumberingLevelDefinition[] = [];
	let active = 0;
	let sourceView: EditorView | undefined;
	const valid = () => {
		const integer = (input: HTMLInputElement) =>
			input.value !== '' &&
			Number.isInteger(Number(input.value)) &&
			Number(input.value) >= 0 &&
			Number(input.value) <= 32767;
		const distance = (input: HTMLInputElement) =>
			input.value !== '' &&
			Number.isFinite(Number(input.value)) &&
			Math.abs(Number(input.value)) <= 22;
		const invalidPlaceholder = [...pattern.value.matchAll(/%([1-9])/g)].some(
			(match) => Number(match[1]) > active + 1,
		);
		const markerValid =
			(markerSize.value === '' ||
				(Number(markerSize.value) >= 1 && Number(markerSize.value) <= 400)) &&
			(markerColor.value.trim() === '' || /^#[0-9a-f]{6}$/i.test(markerColor.value.trim()));
		const tabOk =
			tabStop.value === '' || (Number(tabStop.value) >= 0 && Number(tabStop.value) <= 22);
		const result =
			tabOk &&
			integer(start) &&
			distance(aligned) &&
			distance(indent) &&
			!invalidPlaceholder &&
			markerValid;
		message.textContent = !markerValid
			? 'Enter a font size from 1 to 400 and a colour such as #1F4E79.'
			: result
				? ''
				: 'Enter a start from 0 to 32767, positions within 22 inches and placeholders for this level or its ancestors.';
		localizeElement(message, locale);
		message.hidden = result;
		ok.disabled = !result;
		return result;
	};
	const store = () => {
		if (!valid()) return false;
		const draft = drafts[active]!;
		draft.numFmt = style.value as NumberingLevelDefinition['numFmt'];
		draft.lvlText = pattern.value;
		draft.start = Number(start.value);
		draft.lvlJc = alignment.value as NonNullable<NumberingLevelDefinition['lvlJc']>;
		draft.isLgl = legal.input.checked;
		if (restart.value === 'default') delete draft.lvlRestart;
		else draft.lvlRestart = Number(restart.value);
		if (tabStop.value === '') delete draft.tabStopTwips;
		else draft.tabStopTwips = twips(Math.round(Number(tabStop.value) * 1440));
		if (linkedStyle.value) {
			draft.paragraphStyleId = linkedStyle.value;
			// A style belongs to one level of a list.
			drafts.forEach((other, index) => {
				if (index !== active && other.paragraphStyleId === linkedStyle.value)
					delete other.paragraphStyleId;
			});
		} else delete draft.paragraphStyleId;
		draft.suffix = suffix.value as NonNullable<NumberingLevelDefinition['suffix']>;
		const marker: NonNullable<NumberingLevelDefinition['markerFormat']> = {};
		if (markerFont.value.trim()) marker.fontFamily = markerFont.value.trim();
		if (markerSize.value !== '')
			marker.fontSizeHalfPoints = Math.round(Number(markerSize.value) * 2);
		if (markerColor.value.trim()) marker.color = markerColor.value.trim().toLowerCase();
		if (markerBold.input.checked) marker.bold = true;
		if (markerItalic.input.checked) marker.italic = true;
		if (Object.keys(marker).length) draft.markerFormat = marker;
		else delete draft.markerFormat;
		draft.indentLeftTwips = signedTwips(Math.round(Number(indent.value) * 1440));
		const difference = Math.round((Number(indent.value) - Number(aligned.value)) * 1440);
		delete draft.hangingTwips;
		delete draft.firstLineTwips;
		if (difference >= 0) draft.hangingTwips = twips(difference);
		else draft.firstLineTwips = twips(-difference);
		return true;
	};
	const renderPreview = () => {
		preview.replaceChildren(
			...drafts.map((draft, i) => {
				const line = document.createElement('div');
				const text =
					draft.numFmt === 'bullet'
						? draft.lvlText
						: draft.numFmt === 'none'
							? ''
							: draft.lvlText.replace(/%([1-9])/g, (token, digit: string) => {
									const ancestor = drafts[Number(digit) - 1];
									return ancestor && Number(digit) <= i + 1
										? formatListNumber(draft.isLgl ? 'decimal' : ancestor.numFmt, ancestor.start)
										: token;
								});
				line.textContent = `${text}${draft.suffix === 'none' ? '' : ' '}${translate(locale, 'Level')} ${i + 1}`;
				line.style.marginLeft = `${i * 8}px`;
				line.style.fontWeight = i === active ? 'bold' : 'normal';
				return line;
			}),
		);
	};
	const read = () => {
		const draft = drafts[active]!;
		level.value = String(active);
		if (![...style.options].some((option) => option.value === draft.numFmt))
			style.append(new Option(draft.numFmt, draft.numFmt));
		style.value = draft.numFmt;
		pattern.value = draft.lvlText;
		start.value = String(draft.start);
		alignment.value = draft.lvlJc ?? 'left';
		indent.value = String((draft.indentLeftTwips ?? 0) / 1440);
		aligned.value = String(
			((draft.indentLeftTwips ?? 0) - (draft.hangingTwips ?? 0) + (draft.firstLineTwips ?? 0)) /
				1440,
		);
		suffix.value = draft.suffix ?? 'tab';
		const styles = Object.values(getModel().paragraphStyles?.styles ?? {}).sort((a, b) =>
			(a.name ?? a.id).localeCompare(b.name ?? b.id),
		);
		linkedStyle.replaceChildren(
			new Option('(No style)', ''),
			...styles.map((item) => new Option(item.name ?? item.id, item.id)),
		);
		if (draft.paragraphStyleId && !styles.some((item) => item.id === draft.paragraphStyleId))
			linkedStyle.append(new Option(draft.paragraphStyleId, draft.paragraphStyleId));
		linkedStyle.value = draft.paragraphStyleId ?? '';
		tabStop.value = draft.tabStopTwips !== undefined ? String(draft.tabStopTwips / 1440) : '';
		const marker = draft.markerFormat;
		markerFont.value = marker?.fontFamily ?? '';
		markerSize.value = marker?.fontSizeHalfPoints ? String(marker.fontSizeHalfPoints / 2) : '';
		markerColor.value = marker?.color ?? '';
		markerBold.input.checked = Boolean(marker?.bold);
		markerItalic.input.checked = Boolean(marker?.italic);
		legal.input.checked = Boolean(draft.isLgl);
		restart.replaceChildren(
			new Option('Previous level (default)', 'default'),
			new Option('Never', '0'),
			...Array.from({ length: active }, (_, i) => new Option(`Level ${i + 1}`, String(i + 1))),
		);
		restart.value =
			draft.lvlRestart !== undefined && draft.lvlRestart <= active
				? String(draft.lvlRestart)
				: 'default';
		restart.disabled = active === 0;
		localizeElement(restart, locale);
		valid();
		renderPreview();
	};
	level.addEventListener('change', () => {
		const next = Number(level.value);
		if (!store()) {
			level.value = String(active);
			return;
		}
		active = next;
		read();
	});
	for (const control of [
		style,
		pattern,
		start,
		restart,
		legal.input,
		alignment,
		aligned,
		indent,
		suffix,
		tabStop,
		linkedStyle,
		markerFont,
		markerSize,
		markerColor,
		markerBold.input,
		markerItalic.input,
	])
		control.addEventListener('input', () => {
			if (store()) renderPreview();
		});
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	cancel.addEventListener('click', hide);
	const submit = () => {
		const view = getView();
		if (!view?.editable || view !== sourceView || !canEdit() || !store()) return;
		applyCustomList(view, getModel(), drafts);
		hide();
	};
	ok.addEventListener('click', submit);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			hide();
		} else if (event.key === 'Enter' && event.target instanceof HTMLInputElement) {
			event.preventDefault();
			submit();
		}
	});
	return {
		element,
		close: hide,
		get isOpen() {
			return !element.hidden;
		},
		setLocale(next) {
			locale = next;
		},
		open() {
			const view = getView();
			if (!view?.editable || !canEdit()) return;
			sourceView = view;
			const model = getModel();
			const defaults = ensureListDefinition(undefined, 'multilevel').catalog.abstractNums['0']!
				.levels;
			let paragraph = view.state.selection.$from.parent;
			if (paragraph.type.name !== 'paragraph')
				view.state.doc.nodesBetween(view.state.selection.from, view.state.selection.to, (node) => {
					if (paragraph.type.name !== 'paragraph' && node.type.name === 'paragraph')
						paragraph = node;
				});
			const numbering = nodeNumbering(paragraph, model);
			const numId = numbering?.numId;
			drafts = Array.from({ length: 9 }, (_, i) => ({
				...(model.numberingCatalog && numId
					? (resolveNumberingLevel(model.numberingCatalog, String(numId), i) ?? defaults[i]!)
					: defaults[i]!),
			}));
			active = Math.min(8, Math.max(0, numbering?.level ?? Number(paragraph.attrs.ilvl ?? 0)));
			element.replaceChildren(...content);
			localizeElement(element, locale);
			read();
			element.hidden = false;
			level.focus();
		},
	};
}
