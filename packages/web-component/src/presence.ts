import { getVersion, sendableSteps } from 'prosemirror-collab';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { validId } from './collaboration-protocol';

export const PRESENCE_PALETTE = [
	'#2563eb',
	'#c2410c',
	'#15803d',
	'#a21caf',
	'#0e7490',
	'#b91c1c',
	'#4f46e5',
	'#a16207',
] as const;

export function validatePresenceProfile(profile: { name: string; color: string }) {
	const name = profile.name.trim();
	if (!name || name.length > 80 || /[\u0000-\u001f\u007f]/.test(name))
		throw new Error('Presence name must contain 1 to 80 printable characters');
	if (!PRESENCE_PALETTE.includes(profile.color as (typeof PRESENCE_PALETTE)[number]))
		throw new Error('Presence color must be from the shared palette');
	return { name, color: profile.color };
}

export interface PresenceConfig {
	sessionId: string;
	clientId: string;
}

export interface PresenceMessage {
	protocol: 1;
	sessionId: string;
	clientId: string;
	kind: 'selection' | 'leave';
	sequence: number;
	version: number;
	name?: string;
	color?: string;
	anchor?: number;
	head?: number;
}

export type PresenceReceiveResult =
	| { status: 'applied'; transaction: Transaction }
	| {
			status: 'duplicate' | 'stale' | 'out-of-order' | 'wrong-session' | 'invalid';
			reason?: string;
	  };

interface PeerPresence {
	clientId: string;
	name: string;
	color: (typeof PRESENCE_PALETTE)[number];
	anchor: number;
	head: number;
	fingerprint: string;
	sequence: number;
}

const MAX_PEERS = 100;
const MAX_RECENT = 200;
const presenceKey = new PluginKey<Map<string, PeerPresence>>('docx-presence');

function peerDecorations(doc: EditorState['doc'], peers: Map<string, PeerPresence>): DecorationSet {
	const decorations: Decoration[] = [];
	for (const peer of peers.values()) {
		const from = Math.min(peer.anchor, peer.head);
		const to = Math.max(peer.anchor, peer.head);
		if (from < to) {
			decorations.push(
				Decoration.inline(from, to, {
					class: 'dve-peer-selection',
					style: `background-color: ${peer.color}33; box-shadow: inset 0 -2px ${peer.color}`,
					'aria-label': `${peer.name}'s selection`,
				}),
			);
		}
		const cursor = document.createElement('span');
		cursor.className = 'dve-peer-cursor';
		cursor.style.borderColor = peer.color;
		cursor.style.position = 'relative';
		cursor.style.borderLeftWidth = '2px';
		cursor.style.borderLeftStyle = 'solid';
		cursor.style.marginInline = '-1px';
		cursor.setAttribute('role', 'img');
		cursor.setAttribute('aria-label', `${peer.name}'s cursor`);
		cursor.contentEditable = 'false';
		const label = document.createElement('span');
		label.className = 'dve-peer-cursor-label';
		label.textContent = peer.name;
		label.style.backgroundColor = peer.color;
		label.style.position = 'absolute';
		label.style.left = '-2px';
		label.style.bottom = '100%';
		label.style.padding = '2px 6px';
		label.style.borderRadius = '3px 3px 3px 0';
		label.style.color = '#fff';
		label.style.font = '600 10px/1.4 system-ui, sans-serif';
		label.style.whiteSpace = 'nowrap';
		label.style.pointerEvents = 'none';
		cursor.append(label);
		decorations.push(
			Decoration.widget(peer.head, cursor, {
				key: `peer-${peer.clientId}-${peer.sequence}`,
				side: 1,
			}),
		);
	}
	return DecorationSet.create(doc, decorations);
}

function createPresencePlugin(): Plugin {
	return new Plugin({
		key: presenceKey,
		state: {
			init: () => new Map(),
			apply(transaction, previous) {
				const next = new Map<string, PeerPresence>();
				for (const [id, peer] of previous) {
					const collapsed = peer.anchor === peer.head;
					const anchor = transaction.mapping.map(peer.anchor, -1);
					const head = transaction.mapping.map(peer.head, collapsed ? -1 : 1);
					next.set(id, {
						...peer,
						anchor: Math.max(0, Math.min(anchor, transaction.doc.content.size)),
						head: Math.max(0, Math.min(head, transaction.doc.content.size)),
					});
				}
				const update = transaction.getMeta(presenceKey) as
					| { remove: string }
					| { peer: PeerPresence }
					| undefined;
				if (update && 'remove' in update) next.delete(update.remove);
				if (update && 'peer' in update) next.set(update.peer.clientId, update.peer);
				return next;
			},
		},
		props: {
			decorations(state) {
				return peerDecorations(state.doc, presenceKey.getState(state) ?? new Map());
			},
		},
	});
}

export function getPresenceDecorations(state: EditorState): DecorationSet {
	return peerDecorations(state.doc, presenceKey.getState(state) ?? new Map());
}

/** Transient selection sharing for a host-provided transport; it does not send data itself. */
export class PresenceClient {
	readonly plugin = createPresencePlugin();
	readonly sessionId: string;
	readonly clientId: string;
	private readonly received = new Map<string, { sequence: number; fingerprint: string }>();
	private sequence = 0;

	constructor(config: PresenceConfig) {
		if (!validId(config.sessionId) || !validId(config.clientId))
			throw new Error('Presence sessionId and clientId must be non-empty strings');
		this.sessionId = config.sessionId;
		this.clientId = config.clientId;
	}

	publish(state: EditorState, profile: { name: string; color: string }): PresenceMessage | null {
		const { name, color } = validatePresenceProfile(profile);
		if (sendableSteps(state)?.steps.length) return null;
		return {
			protocol: 1,
			sessionId: this.sessionId,
			clientId: this.clientId,
			kind: 'selection',
			sequence: ++this.sequence,
			version: versionOf(state),
			name,
			color,
			anchor: state.selection.anchor,
			head: state.selection.head,
		};
	}

	leave(state: EditorState): PresenceMessage {
		return {
			protocol: 1,
			sessionId: this.sessionId,
			clientId: this.clientId,
			kind: 'leave',
			sequence: ++this.sequence,
			version: versionOf(state),
		};
	}

	receive(state: EditorState, input: unknown): PresenceReceiveResult {
		const parsed = parsePresence(input, this.sessionId, state);
		if ('status' in parsed) return parsed;
		if (parsed.clientId === this.clientId) return { status: 'duplicate' };
		const previousMessage = this.received.get(parsed.clientId);
		if (previousMessage && parsed.sequence < previousMessage.sequence) return { status: 'stale' };
		if (previousMessage && parsed.sequence === previousMessage.sequence)
			return previousMessage.fingerprint === parsed.fingerprint
				? { status: 'duplicate' }
				: { status: 'invalid', reason: 'Presence sequence was reused' };
		const previous = presenceKey.getState(state)?.get(parsed.clientId);
		if (
			parsed.kind === 'selection' &&
			!previous &&
			(presenceKey.getState(state)?.size ?? 0) >= MAX_PEERS
		)
			return { status: 'invalid', reason: 'Presence peer limit reached' };
		const transaction = state.tr.setMeta(
			presenceKey,
			parsed.kind === 'leave' ? { remove: parsed.clientId } : { peer: parsed },
		);
		this.received.set(parsed.clientId, {
			sequence: parsed.sequence,
			fingerprint: parsed.fingerprint,
		});
		if (this.received.size > MAX_RECENT) this.received.delete(this.received.keys().next().value!);
		return { status: 'applied', transaction };
	}
}

function versionOf(state: EditorState): number {
	try {
		return getVersion(state);
	} catch {
		throw new Error('EditorState is missing the collaboration plugin');
	}
}

function parsePresence(
	input: unknown,
	sessionId: string,
	state: EditorState,
):
	| (PeerPresence & { kind: 'selection' })
	| { kind: 'leave'; clientId: string; sequence: number; fingerprint: string }
	| { status: 'wrong-session' | 'stale' | 'out-of-order' | 'invalid'; reason?: string } {
	if (!input || typeof input !== 'object' || Array.isArray(input))
		return { status: 'invalid', reason: 'Presence message must be an object' };
	const message = input as Record<string, unknown>;
	if (message.protocol !== 1 || !validId(message.clientId))
		return { status: 'invalid', reason: 'Invalid presence protocol or client ID' };
	if (message.sessionId !== sessionId) return { status: 'wrong-session' };
	if (message.kind !== 'selection' && message.kind !== 'leave')
		return { status: 'invalid', reason: 'Invalid presence message kind' };
	if (!Number.isSafeInteger(message.sequence) || Number(message.sequence) < 1)
		return { status: 'invalid', reason: 'Invalid presence sequence' };
	if (!Number.isSafeInteger(message.version) || Number(message.version) < 0)
		return { status: 'invalid', reason: 'Invalid presence version' };
	const current = versionOf(state);
	if (message.kind !== 'leave' && Number(message.version) < current) return { status: 'stale' };
	if (message.kind !== 'leave' && Number(message.version) > current)
		return { status: 'out-of-order' };
	if (message.kind === 'leave')
		return {
			kind: 'leave',
			clientId: message.clientId,
			sequence: Number(message.sequence),
			fingerprint: JSON.stringify(message),
		};
	const name = typeof message.name === 'string' ? message.name.trim() : '';
	if (!name || name.length > 80 || /[\u0000-\u001f\u007f]/.test(name))
		return { status: 'invalid', reason: 'Invalid peer name' };
	if (!PRESENCE_PALETTE.includes(message.color as (typeof PRESENCE_PALETTE)[number]))
		return { status: 'invalid', reason: 'Color must be from the shared palette' };
	if (
		!Number.isSafeInteger(message.anchor) ||
		!Number.isSafeInteger(message.head) ||
		Number(message.anchor) < 0 ||
		Number(message.head) < 0 ||
		Number(message.anchor) > state.doc.content.size ||
		Number(message.head) > state.doc.content.size
	)
		return { status: 'invalid', reason: 'Selection is outside the current document' };
	return {
		kind: 'selection',
		sequence: Number(message.sequence),
		clientId: message.clientId,
		name,
		color: message.color as (typeof PRESENCE_PALETTE)[number],
		anchor: Number(message.anchor),
		head: Number(message.head),
		fingerprint: JSON.stringify(message),
	};
}
