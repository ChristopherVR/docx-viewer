import type { Comment, DocumentModel } from '@christophervr/docx-core';
import { localizeElement, normalizeEditorLocale, translate, type EditorLocale } from './localization';

export interface CommentsPanelOptions {
	getModel: () => DocumentModel;
	canAdd: () => boolean;
	onAdd: (text: string) => void;
	onReply: (parentId: string, text: string) => void;
	onResolve: (id: string, resolved: boolean) => void;
	onDelete: (id: string) => void;
	onClose?: () => void;
}
export interface CommentsPanelHandle {
	element: HTMLElement;
	open(): void;
	close(): void;
	refresh(): void;
	setLocale(locale: string): void;
	readonly isOpen: boolean;
}

const styleText = `
	.dve-comments-panel{display:flex;flex-direction:column;gap:8px;padding:10px 12px;border-left:1px solid var(--line,#ddd);background:var(--surface,#fff);color:var(--ink,#222);font:12px/1.4 'Segoe UI',Arial,sans-serif;width:clamp(220px,26vw,320px);overflow-y:auto}
	.dve-comments-panel[hidden]{display:none}
	.dve-comments-panel h2{margin:0;font-size:13px}
	.dve-comment-new textarea,.dve-comment-reply textarea{box-sizing:border-box;width:100%;min-height:44px;resize:vertical;border:1px solid var(--line,#ddd);border-radius:3px;padding:5px 7px;font:inherit;background:var(--surface,#fff);color:var(--ink,#222)}
	.dve-comment-new button,.dve-comment-actions button,.dve-comment-reply button{border:1px solid var(--line,#ddd);border-radius:3px;background:var(--surface,#fff);color:var(--ink,#222);font:inherit;cursor:pointer;padding:3px 8px}
	.dve-comment-new button:disabled,.dve-comment-actions button:disabled{opacity:.5;cursor:default}
	.dve-comment-thread{border:1px solid var(--line,#ddd);border-radius:4px;padding:7px 8px;display:flex;flex-direction:column;gap:4px}
	.dve-comment-thread[data-resolved="true"]{opacity:.65}
	.dve-comment-meta{display:flex;justify-content:space-between;gap:6px;color:var(--muted,#666);font-size:11px}
	.dve-comment-text{white-space:pre-wrap}
	.dve-comment-actions{display:flex;gap:5px;flex-wrap:wrap}
	.dve-comment-reply{margin-left:12px;display:flex;flex-direction:column;gap:4px}
	.dve-comment-empty{color:var(--muted,#666)}
`;

function row(comment: Comment, indent: boolean, locale: EditorLocale, options: CommentsPanelOptions): HTMLElement {
	const thread = document.createElement('div');
	thread.className = 'dve-comment-thread';
	if (indent) thread.style.marginLeft = '14px';
	thread.dataset.resolved = String(Boolean(comment.resolved));
	const meta = document.createElement('div');
	meta.className = 'dve-comment-meta';
	meta.append(
		Object.assign(document.createElement('span'), { textContent: comment.author }),
		Object.assign(document.createElement('span'), { textContent: comment.date ?? '' }),
	);
	const text = document.createElement('p');
	text.className = 'dve-comment-text';
	text.textContent = comment.text;
	const actions = document.createElement('div');
	actions.className = 'dve-comment-actions';
	if (!comment.parentId) {
		const resolve = document.createElement('button');
		resolve.type = 'button';
		resolve.textContent = translate(locale, comment.resolved ? 'Reopen' : 'Resolve');
		resolve.addEventListener('click', () => options.onResolve(comment.id, !comment.resolved));
		actions.append(resolve);
	}
	const del = document.createElement('button');
	del.type = 'button';
	del.textContent = translate(locale, 'Delete');
	del.addEventListener('click', () => options.onDelete(comment.id));
	actions.append(del);
	thread.append(meta, text, actions);
	if (!comment.parentId) {
		const replyForm = document.createElement('div');
		replyForm.className = 'dve-comment-reply';
		const input = document.createElement('textarea');
		input.setAttribute('aria-label', translate(locale, 'Reply'));
		input.placeholder = translate(locale, 'Reply');
		const send = document.createElement('button');
		send.type = 'button';
		send.textContent = translate(locale, 'Reply');
		send.addEventListener('click', () => {
			if (!input.value.trim()) return;
			options.onReply(comment.id, input.value.trim());
			input.value = '';
		});
		replyForm.append(input, send);
		thread.append(replyForm);
	}
	return thread;
}

export function createCommentsPanel(options: CommentsPanelOptions): CommentsPanelHandle {
	let locale: EditorLocale = 'en';
	const panel = document.createElement('aside');
	panel.className = 'dve-comments-panel';
	panel.setAttribute('role', 'complementary');
	panel.setAttribute('aria-label', 'Comments');
	panel.hidden = true;
	const style = document.createElement('style');
	style.textContent = styleText;
	const heading = document.createElement('h2');
	heading.textContent = translate(locale, 'Comments');
	const closeButton = document.createElement('button');
	closeButton.type = 'button';
	closeButton.textContent = translate(locale, 'Close comments');
	closeButton.setAttribute('aria-label', 'Close comments');
	closeButton.addEventListener('click', () => close());
	const newForm = document.createElement('div');
	newForm.className = 'dve-comment-new';
	const newInput = document.createElement('textarea');
	newInput.setAttribute('aria-label', 'New comment');
	newInput.placeholder = translate(locale, 'New comment');
	const addButton = document.createElement('button');
	addButton.type = 'button';
	addButton.textContent = translate(locale, 'Add comment');
	addButton.addEventListener('click', () => {
		if (!newInput.value.trim()) return;
		options.onAdd(newInput.value.trim());
		newInput.value = '';
		refresh();
	});
	newForm.append(newInput, addButton);
	const list = document.createElement('div');
	panel.append(style, heading, closeButton, newForm, list);

	let isOpen = false;
	function close() {
		isOpen = false;
		panel.hidden = true;
		options.onClose?.();
	}
	function refresh() {
		addButton.disabled = !options.canAdd();
		const comments = options.getModel().comments ?? [];
		list.replaceChildren();
		const top = comments.filter((comment) => !comment.parentId);
		if (!top.length) {
			const empty = document.createElement('p');
			empty.className = 'dve-comment-empty';
			empty.textContent = translate(locale, 'No comments');
			list.append(empty);
			return;
		}
		for (const comment of top) {
			list.append(row(comment, false, locale, options));
			for (const reply of comments.filter((item) => item.parentId === comment.id))
				list.append(row(reply, true, locale, options));
		}
	}
	localizeElement(panel, locale);
	refresh();
	return {
		element: panel,
		get isOpen() {
			return isOpen;
		},
		open() {
			isOpen = true;
			panel.hidden = false;
			refresh();
		},
		close,
		refresh,
		setLocale(value: string) {
			locale = normalizeEditorLocale(value);
			heading.textContent = translate(locale, 'Comments');
			newInput.placeholder = translate(locale, 'New comment');
			addButton.textContent = translate(locale, 'Add comment');
			refresh();
		},
	};
}
