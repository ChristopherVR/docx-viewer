import { collab, getVersion, receiveTransaction, sendableSteps } from 'prosemirror-collab';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import { IdempotencyCache, classifyVersion, isValidId } from 'ooxml-core/collab';
import {
	freezeBatch,
	parseBatch,
	type CollaborationConfig,
	type StepBatch,
} from './collaboration-protocol';
export type { CollaborationConfig, StepBatch } from './collaboration-protocol';
export { CollaborationAuthority } from './collaboration-authority';
export type { CollaborationAuthorityConfig, AuthorityResult } from './collaboration-authority';

export type ClientReceiveResult =
	| { status: 'applied'; transaction: Transaction }
	| {
			status: 'duplicate' | 'stale' | 'out-of-order' | 'wrong-session' | 'invalid';
			reason?: string;
	  };

const MAX_RECENT_BATCHES = 1_000;

function fromSendable(
	data: NonNullable<ReturnType<typeof sendableSteps>>,
	config: CollaborationConfig,
	batchId: string,
): StepBatch {
	return freezeBatch({
		protocol: 1,
		sessionId: config.sessionId,
		batchId,
		version: data.version,
		clientId: config.clientId,
		steps: data.steps.map((step) => step.toJSON()),
	});
}

/** Client-side plugin and transport helpers. It does not open a network connection. */
export class CollaborationClient {
	readonly plugin: Plugin;
	readonly sessionId: string;
	readonly clientId: string;
	private readonly config: CollaborationConfig;
	private sequence = 0;
	private pending?: StepBatch | undefined;
	private readonly received = new IdempotencyCache(MAX_RECENT_BATCHES);

	constructor(config: CollaborationConfig) {
		if (!isValidId(config.sessionId) || !isValidId(config.clientId))
			throw new Error('Collaboration sessionId and clientId must be non-empty strings');
		const version = config.version ?? 0;
		if (!Number.isSafeInteger(version) || version < 0) throw new Error('Invalid starting version');
		this.config = { ...config, version };
		this.sessionId = config.sessionId;
		this.clientId = config.clientId;
		this.plugin = collab({ clientID: config.clientId, version });
	}

	get hasPendingAck(): boolean {
		return Boolean(this.pending);
	}

	pendingStepCount(state: EditorState): number {
		return sendableSteps(state)?.steps.length ?? 0;
	}

	/** Returns the same envelope until acknowledged, so transport retries are idempotent. */
	createPendingBatch(state: EditorState): StepBatch | null {
		if (this.pending) return this.pending;
		const sendable = sendableSteps(state);
		if (!sendable) return null;
		if (String(sendable.clientID) !== this.clientId)
			throw new Error('EditorState uses a different collaboration client ID');
		// Batch IDs are scoped by clientId in the protocol cache, so keep this
		// identifier bounded even when clientId itself uses the full 160 chars.
		this.pending = fromSendable(sendable, this.config, String(++this.sequence));
		return this.pending;
	}

	/** Converts one ordered authority batch into the transaction hosts should dispatch. */
	receive(state: EditorState, input: unknown): ClientReceiveResult {
		const parsed = parseBatch(input, state.schema);
		if (typeof parsed === 'string') return { status: 'invalid', reason: parsed };
		const { batch, steps, fingerprint } = parsed;
		if (batch.sessionId !== this.sessionId) return { status: 'wrong-session' };
		const seen = this.received.check(batch.clientId, batch.batchId, fingerprint);
		if (seen.status === 'duplicate') return { status: 'duplicate' };
		if (seen.status === 'reused') return { status: 'invalid' };
		const order = classifyVersion(batch.version, collaborationVersion(state));
		if (order === 'stale' || order === 'out-of-order') return { status: order };
		if (batch.clientId === this.clientId) {
			const sendable = sendableSteps(state);
			const acknowledged = JSON.stringify(steps.map((step) => step.toJSON()));
			if (
				!this.pending ||
				this.pending.batchId !== batch.batchId ||
				!sendable ||
				sendable.steps.length < steps.length ||
				JSON.stringify(sendable.steps.slice(0, steps.length).map((step) => step.toJSON())) !==
					acknowledged
			)
				return { status: 'invalid', reason: 'Unexpected or mismatched acknowledgement' };
		}
		let transaction: Transaction;
		try {
			transaction = receiveTransaction(
				state,
				steps,
				steps.map(() => batch.clientId),
				{ mapSelectionBackward: true },
			);
		} catch (cause) {
			return {
				status: 'invalid',
				reason: cause instanceof Error ? cause.message : 'Could not apply collaboration steps',
			};
		}
		this.received.remember(batch.clientId, batch.batchId, fingerprint, undefined);
		if (batch.clientId === this.clientId && this.pending?.batchId === batch.batchId)
			this.pending = undefined;
		return { status: 'applied', transaction };
	}
}

function collaborationVersion(state: EditorState): number {
	try {
		return getVersion(state);
	} catch {
		throw new Error('EditorState is missing the collaboration plugin');
	}
}
