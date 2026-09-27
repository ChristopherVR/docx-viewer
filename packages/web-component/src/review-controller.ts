import type { EditorView } from 'prosemirror-view';
import type { Comment, DocumentModel } from '@christophervr/docx-core';
import { addComment, deleteComment, removeCommentAnchor, replyToComment, resolveComment } from './comment-commands';
import { createCommentsPanel, type CommentsPanelHandle } from './comments-panel';
import {
	acceptAllChanges,
	acceptChangeAtCursor,
	goToNextChange,
	goToPreviousChange,
	rejectAllChanges,
	rejectChangeAtCursor,
} from './review-commands';
import type { RibbonAction } from './ribbon';

/** Narrow view of the host element `ReviewController` needs; keeps component.ts's own diff small. */
export interface ReviewHost {
	getModel(): DocumentModel;
	setModel(model: DocumentModel): void;
	getView(): EditorView | undefined;
	getReviewAuthor(): string;
	getCollaborationIds(): ((kind: string) => string) | undefined;
	notifyChange(): void;
	refresh(): void;
}

/** Owns comment add/reply/resolve/delete and tracked-change accept/reject/navigate/toggle commands. */
export class ReviewController {
	readonly commentsPanel: CommentsPanelHandle;
	constructor(private host: ReviewHost) {
		this.commentsPanel = createCommentsPanel({
			getModel: () => host.getModel(),
			canAdd: () => {
				const view = host.getView();
				return Boolean(view && !view.state.selection.empty);
			},
			onAdd: (text) => this.addComment(text),
			onReply: (parentId, text) =>
				this.updateComments(
					replyToComment(host.getModel().comments ?? [], parentId, host.getReviewAuthor(), text, host.getCollaborationIds()),
				),
			onResolve: (id, resolved) =>
				this.updateComments(resolveComment(host.getModel().comments ?? [], id, resolved)),
			onDelete: (id) => {
				const view = host.getView();
				if (view) removeCommentAnchor(view, id);
				this.updateComments(deleteComment(host.getModel().comments ?? [], id));
			},
			onClose: () => host.getView()?.focus(),
		});
	}
	get commentsOpen(): boolean {
		return this.commentsPanel.isOpen;
	}
	setLocale(locale: string): void {
		this.commentsPanel.setLocale(locale);
	}

	private updateComments(next: Comment[]) {
		this.host.setModel({ ...this.host.getModel(), comments: next });
		this.host.notifyChange();
		this.host.refresh();
	}
	private addComment(text: string) {
		const view = this.host.getView();
		if (!view) return;
		const comment = addComment(view, this.host.getReviewAuthor(), text, this.host.getCollaborationIds());
		if (comment) this.updateComments([...(this.host.getModel().comments ?? []), comment]);
	}

	handleReview(key: Extract<RibbonAction, { type: 'review' }>['key']): void {
		if (key === 'trackChanges') {
			this.host.setModel({ ...this.host.getModel(), trackChanges: !this.host.getModel().trackChanges });
			this.host.notifyChange();
			this.host.refresh();
			return;
		}
		const view = this.host.getView();
		if (!view) return;
		const byKey: Record<string, (view: EditorView) => boolean> = {
			acceptOne: acceptChangeAtCursor,
			rejectOne: rejectChangeAtCursor,
			acceptAll: acceptAllChanges,
			rejectAll: rejectAllChanges,
			previous: goToPreviousChange,
			next: goToNextChange,
		};
		byKey[key]?.(view);
	}
	handleComments(key: 'toggle' | 'add'): void {
		if (key === 'toggle') {
			if (this.commentsPanel.isOpen) this.commentsPanel.close();
			else this.commentsPanel.open();
		} else this.commentsPanel.open();
		this.host.refresh();
	}
}
