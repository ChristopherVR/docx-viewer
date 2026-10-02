import type { DocumentModel } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { goToTarget, goToTargets, type GoToKind } from './go-to';
import { localeOf, translate } from './localization';
import { mountPopover } from './ribbon-popover';

const KINDS: Array<[GoToKind, 'Heading' | 'Bookmark' | 'Table' | 'Graphic']> = [
	['heading', 'Heading'],
	['bookmark', 'Bookmark'],
	['table', 'Table'],
	['graphic', 'Graphic'],
];

/** Word's Go To, as a panel under the Find menu: pick what to go to, then pick the item. */
export function openGoToPanel(
	anchor: HTMLElement,
	view: EditorView,
	model: DocumentModel,
	onClose: () => void,
): void {
	const locale = localeOf(anchor);
	const pop = document.createElement('div');
	pop.className = 'ribbon-popover go-to-panel';
	pop.setAttribute('role', 'dialog');
	pop.setAttribute('aria-label', translate(locale, 'Go to'));
	const label = document.createElement('label');
	const text = document.createElement('span');
	text.textContent = translate(locale, 'Go to what');
	const kind = document.createElement('select');
	kind.setAttribute('aria-label', translate(locale, 'Go to what'));
	for (const [value, name] of KINDS) kind.append(new Option(translate(locale, name), value));
	label.append(text, kind);
	const list = document.createElement('div');
	list.className = 'go-to-list';
	list.setAttribute('role', 'listbox');
	const fill = () => {
		const targets = goToTargets(view.state.doc, model, kind.value as GoToKind);
		if (!targets.length) {
			const empty = document.createElement('p');
			empty.textContent = translate(locale, 'Nothing to go to.');
			list.replaceChildren(empty);
			return;
		}
		list.replaceChildren(
			...targets.map((target) => {
				const item = document.createElement('button');
				item.type = 'button';
				item.setAttribute('role', 'option');
				item.textContent = target.label;
				item.style.whiteSpace = 'pre';
				item.addEventListener('mousedown', (event) => event.preventDefault());
				item.addEventListener('click', () => {
					goToTarget(view, target);
					onClose();
				});
				return item;
			}),
		);
	};
	kind.addEventListener('change', fill);
	pop.append(label, list);
	fill();
	if (mountPopover(anchor, pop, anchor)) kind.focus();
}
