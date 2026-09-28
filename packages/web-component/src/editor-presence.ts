import { emit } from './events';
import type { EditorView } from 'prosemirror-view';
import { sendableSteps } from 'prosemirror-collab';
import {
	PresenceClient,
	validatePresenceProfile,
	type PresenceConfig,
	type PresenceMessage,
	type PresenceReceiveResult,
} from './presence';

export class EditorPresence {
	readonly client: PresenceClient;
	private profile?: { name: string; color: string };
	private scheduled = false;
	constructor(
		config: PresenceConfig,
		private element: HTMLElement,
		private view: () => EditorView | undefined,
	) {
		this.client = new PresenceClient(config);
	}
	publish(profile: { name: string; color: string }): PresenceMessage | null {
		this.profile = validatePresenceProfile(profile);
		const state = this.view()?.state;
		if (!state || sendableSteps(state)) return null;
		const message = this.client.publish(state, profile);
		if (message) this.emit(message);
		return message;
	}
	receive(input: unknown): PresenceReceiveResult['status'] {
		const view = this.view();
		if (!view) throw new Error('Mount the collaboration editor before receiving presence.');
		// Positions are meaningful only against the acknowledged document version.
		if (
			sendableSteps(view.state) &&
			!(input && typeof input === 'object' && 'kind' in input && input.kind === 'leave')
		)
			return 'out-of-order';
		const result = this.client.receive(view.state, input);
		if (result.status === 'applied') view.dispatch(result.transaction);
		return result.status;
	}
	leave(): PresenceMessage | null {
		this.profile = undefined;
		const state = this.view()?.state;
		if (!state) return null;
		const message = this.client.leave(state);
		this.emit(message);
		return message;
	}
	schedule() {
		if (this.scheduled || !this.profile) return;
		this.scheduled = true;
		queueMicrotask(() => {
			this.scheduled = false;
			if (this.profile && this.element.isConnected) this.publish(this.profile);
		});
	}
	private emit(message: PresenceMessage) {
		emit(this.element, 'presence-send', message);
	}
}
