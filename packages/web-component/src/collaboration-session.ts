import type { EditorView } from 'prosemirror-view';
import { CollaborationClient, type CollaborationConfig, type StepBatch } from './collaboration';
import { createCollaborationIdGenerator } from './collaboration-identity';
import { EditorPresence } from './editor-presence';

/** One editor's collaboration state: step client, presence and collaborative id generator. */
export class CollaborationSession {
	client?: CollaborationClient;
	presence?: EditorPresence;
	ids?: (kind: string) => string;
	private sendScheduled = false;

	constructor(
		private readonly element: HTMLElement,
		private readonly getView: () => EditorView | undefined,
	) {}

	get active(): boolean {
		return Boolean(this.client);
	}

	start(config: CollaborationConfig): void {
		if (this.client) throw new Error('Stop the current collaboration session first.');
		this.client = new CollaborationClient(config);
		this.presence = new EditorPresence(config, this.element, this.getView);
		this.ids = createCollaborationIdGenerator(config.clientId);
	}

	stop(): void {
		this.client = undefined;
		this.presence?.leave();
		this.presence = undefined;
		this.ids = undefined;
	}

	/** Coalesces pending steps into one `collaboration-send` event per microtask. */
	scheduleSend(pending: () => StepBatch | null): void {
		if (!this.client || this.sendScheduled) return;
		this.sendScheduled = true;
		queueMicrotask(() => {
			this.sendScheduled = false;
			if (!this.element.isConnected) return;
			const batch = pending();
			if (batch)
				this.element.dispatchEvent(
					new CustomEvent('collaboration-send', { detail: batch, bubbles: true, composed: true }),
				);
		});
	}
}
