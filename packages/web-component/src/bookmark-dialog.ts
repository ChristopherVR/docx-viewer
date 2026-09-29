import type { EditorView } from 'prosemirror-view';
import {
	addBookmark,
	bookmarkNameProblem,
	deleteBookmark,
	goToBookmark,
	listBookmarks,
	MAX_BOOKMARK_LENGTH,
	type BookmarkNameProblem,
} from './bookmark-commands';
import { dialogButton, labelled, textInput } from './dialog-fields';
import { focusView } from './focus-view';
import type { FormatDialog } from './font-dialog';
import { localizeElement, type EditorLocale } from './localization';

const MESSAGES: Record<BookmarkNameProblem, string> = {
	empty: 'Type a name for the bookmark.',
	start: 'A bookmark name must begin with a letter.',
	characters: 'A bookmark name can contain only letters, numbers and underscores.',
	length: `A bookmark name can be up to ${MAX_BOOKMARK_LENGTH} characters.`,
};

/** Word's Bookmark dialog: name a bookmark for the current paragraph, or go to or delete one. */
export function createBookmarkDialog(getView: () => EditorView | undefined): FormatDialog {
	const element = document.createElement('section');
	element.className = 'dve-dialog dve-format-dialog dve-bookmark-dialog';
	element.setAttribute('role', 'dialog');
	element.setAttribute('aria-label', 'Bookmark');
	element.hidden = true;
	const heading = document.createElement('h2');
	heading.textContent = 'Bookmark';
	const name = textInput();
	const list = document.createElement('select');
	list.size = 6;
	const message = document.createElement('p');
	message.className = 'dve-dialog-message';
	message.setAttribute('role', 'status');
	const add = dialogButton('Add', true);
	const remove = dialogButton('Delete');
	const goTo = dialogButton('Go to');
	const close = dialogButton('Close');
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.append(remove, goTo, close, add);
	element.append(
		heading,
		labelled('Bookmark name', name),
		labelled('Bookmarks', list),
		message,
		actions,
	);
	let locale: EditorLocale = 'en';
	const content = [...element.childNodes];
	element.replaceChildren();

	const say = (text: string) => {
		message.textContent = text;
		localizeElement(message, locale);
	};
	const refresh = () => {
		const view = getView();
		const names = view ? listBookmarks(view) : [];
		list.replaceChildren(
			...[...names].sort((a, b) => a.localeCompare(b)).map((n) => new Option(n, n)),
		);
		const selected = names.includes(name.value) ? name.value : '';
		list.value = selected;
		const problem = bookmarkNameProblem(name.value);
		add.disabled = Boolean(problem) || !view?.editable;
		remove.disabled = !selected || !view?.editable;
		goTo.disabled = !selected;
	};
	const hide = () => {
		element.hidden = true;
		element.replaceChildren();
		focusView(getView());
	};
	name.addEventListener('input', () => {
		const problem = bookmarkNameProblem(name.value);
		say(problem && name.value ? MESSAGES[problem] : '');
		refresh();
	});
	list.addEventListener('change', () => {
		name.value = list.value;
		say('');
		refresh();
	});
	add.addEventListener('click', () => {
		const view = getView();
		if (!view) return;
		const result = addBookmark(view, name.value);
		if (result === 'added' || result === 'moved') hide();
		else if (result !== 'unavailable') say(MESSAGES[result]);
	});
	remove.addEventListener('click', () => {
		const view = getView();
		if (view && deleteBookmark(view, list.value)) {
			name.value = '';
			say('');
			refresh();
		}
	});
	goTo.addEventListener('click', () => {
		const view = getView();
		if (view) goToBookmark(view, list.value);
	});
	close.addEventListener('click', hide);
	element.addEventListener('keydown', (event) => {
		if (event.key === 'Escape') hide();
		else if (event.key === 'Enter' && event.target === name && !add.disabled) {
			event.preventDefault();
			add.click();
		}
	});
	return {
		element,
		open() {
			const view = getView();
			if (!view) return;
			element.replaceChildren(...content);
			localizeElement(element, locale);
			name.value = '';
			say('');
			refresh();
			element.hidden = false;
			name.focus();
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
