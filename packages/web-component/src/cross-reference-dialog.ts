import type { DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import {
	insertCrossReference,
	referenceTargets,
	type ReferenceContent,
	type ReferenceKind,
	type ReferenceTarget,
} from './cross-reference-commands';
import { dialogButton, labelled, selectOf } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';

const KINDS: Array<[ReferenceKind, string]> = [
	['heading', 'Heading'],
	['bookmark', 'Bookmark'],
	['figure', 'Figure'],
	['table', 'Table'],
	['equation', 'Equation'],
];

/** What "Insert reference to" offers for each kind of target. */
function contents(kind: ReferenceKind): Array<[ReferenceContent, string]> {
	if (kind === 'heading')
		return [
			['text', 'Heading text'],
			['page', 'Page number'],
		];
	if (kind === 'bookmark')
		return [
			['text', 'Bookmark text'],
			['page', 'Page number'],
		];
	return [
		['text', 'Entire caption'],
		['label', 'Only label and number'],
		['page', 'Page number'],
	];
}

/** Word's Cross-reference dialog for headings, bookmarks and numbered captions. */
export function createCrossReferenceDialog(
	getView: () => EditorView | undefined,
	getModel: () => DocumentModel,
	pageOf: (id: string) => string | undefined,
): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-cross-reference-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Cross-reference');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Cross-reference';
	const kind = selectOf(KINDS);
	const content = selectOf(contents('heading'));
	const list = document.createElement('select');
	list.size = 8;
	const insert = dialogButton('Insert', true);
	const close = dialogButton('Close');
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(close, insert);
	element.append(
		heading,
		labelled('Reference type', kind),
		labelled('Insert reference to', content),
		labelled('For which item', list),
		actions,
	);
	let locale: EditorLocale = 'en';
	const built = [...element.childNodes];
	element.replaceChildren();
	let targets: ReferenceTarget[] = [];
	const fill = () => {
		const view = getView();
		targets = view ? referenceTargets(view.state.doc, getModel(), kind.value as ReferenceKind) : [];
		list.replaceChildren(
			...targets.map((target, index) => new Option(target.label, String(index))),
		);
		content.replaceChildren(
			...contents(kind.value as ReferenceKind).map(([value, label]) => new Option(label, value)),
		);
		localizeElement(content, locale);
		insert.disabled = !targets.length;
		if (targets.length) list.selectedIndex = 0;
	};
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	kind.addEventListener('change', fill);
	insert.addEventListener('click', () => {
		const view = getView();
		const target = targets[Number(list.value)];
		if (view && target)
			insertCrossReference(view, target, content.value as ReferenceContent, pageOf);
	});
	close.addEventListener('click', hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') hide();
	});
	return {
		element,
		open() {
			if (!getView()?.editable) return;
			element.replaceChildren(...built);
			localizeElement(element, locale);
			kind.value = 'heading';
			fill();
			element.hidden = false;
			kind.focus();
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
