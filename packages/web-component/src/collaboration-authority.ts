import type { Node as ProseMirrorNode, Schema } from 'prosemirror-model';
import { Step, Transform } from 'prosemirror-transform';
import { expectDefined } from './defined';
import { IdempotencyCache, classifyVersion, isValidId } from '@christophervr/ooxml-core/collab';
import { freezeBatch, parseBatch, type StepBatch } from './collaboration-protocol';

export interface CollaborationAuthorityConfig {
	sessionId: string;
	doc: ProseMirrorNode;
	version?: number;
	stepHistoryLimit?: number;
}

export type AuthorityResult =
	| { status: 'accepted' | 'duplicate'; batch: StepBatch }
	| { status: 'rejected'; reason: string };

type RecordStep = { version: number; clientId: string; step: Step };
const MAX_CACHED_REQUESTS = 5_000;

/** ProseMirror's own collab rebase uses this hook, but `prosemirror-transform` does not type it. */
interface MirroringMapping {
	setMirror(from: number, to: number): void;
}
function hasMirror(mapping: object): mapping is MirroringMapping {
	return typeof (mapping as Partial<MirroringMapping>).setMirror === 'function';
}
function setMirror(transform: Transform, from: number, to: number) {
	if (!hasMirror(transform.mapping))
		throw new Error('prosemirror-transform Mapping.setMirror is unavailable.');
	transform.mapping.setMirror(from, to);
}

/** Small in-memory reference authority. A host should replace it with its server and persistence. */
export class CollaborationAuthority {
	readonly sessionId: string;
	private readonly schema: Schema;
	private readonly stepHistoryLimit: number;
	private version: number;
	private documents = new Map<number, ProseMirrorNode>();
	private records: RecordStep[] = [];
	private readonly requests = new IdempotencyCache<StepBatch>(MAX_CACHED_REQUESTS);

	constructor(config: CollaborationAuthorityConfig) {
		if (!isValidId(config.sessionId)) throw new Error('Collaboration sessionId must be non-empty');
		const version = config.version ?? 0;
		if (!Number.isSafeInteger(version) || version < 0) throw new Error('Invalid starting version');
		this.sessionId = config.sessionId;
		this.schema = config.doc.type.schema;
		this.version = version;
		this.stepHistoryLimit = config.stepHistoryLimit ?? 2000;
		if (!Number.isSafeInteger(this.stepHistoryLimit) || this.stepHistoryLimit < 1)
			throw new Error('stepHistoryLimit must be positive');
		this.documents.set(version, config.doc);
	}

	get currentVersion(): number {
		return this.version;
	}

	get doc(): ProseMirrorNode {
		return this.documents.get(this.version)!;
	}

	submit(input: unknown): AuthorityResult {
		const parsed = parseBatch(input, this.schema);
		if (typeof parsed === 'string') return { status: 'rejected', reason: parsed };
		const { batch, steps, fingerprint } = parsed;
		if (batch.sessionId !== this.sessionId) return { status: 'rejected', reason: 'wrong-session' };
		const prior = this.requests.check(batch.clientId, batch.batchId, fingerprint);
		if (prior.status === 'duplicate') return { status: 'duplicate', batch: prior.value };
		if (prior.status === 'reused') return { status: 'rejected', reason: 'batch-id-reused' };
		if (classifyVersion(batch.version, this.version) === 'out-of-order')
			return { status: 'rejected', reason: 'out-of-order' };
		const oldestVersion = this.records[0]?.version ?? this.version;
		if (batch.version < oldestVersion || !this.documents.has(batch.version))
			return { status: 'rejected', reason: 'stale' };

		const baseDoc = this.documents.get(batch.version)!;
		const committed: Step[] = [];
		let transform: Transform;
		try {
			const local = new Transform(baseDoc);
			for (const step of steps) {
				const result = local.maybeStep(step);
				if (result.failed) return { status: 'rejected', reason: `step-failed: ${result.failed}` };
			}

			transform = new Transform(local.doc);
			for (let index = steps.length - 1; index >= 0; index--) {
				const step = expectDefined(steps[index], 'local step');
				const result = transform.maybeStep(
					step.invert(expectDefined(local.docs[index], 'document before local step')),
				);
				if (result.failed) return { status: 'rejected', reason: 'could-not-rebase-local-steps' };
			}
			for (const record of this.records.filter((item) => item.version >= batch.version)) {
				const result = transform.maybeStep(record.step);
				if (result.failed) return { status: 'rejected', reason: 'authority-history-invalid' };
			}

			let mapFrom = steps.length;
			for (const step of steps) {
				const mapped = step.map(transform.mapping.slice(mapFrom));
				mapFrom--;
				if (!mapped) return { status: 'rejected', reason: 'rebase-conflict' };
				const result = transform.maybeStep(mapped);
				if (result.failed) return { status: 'rejected', reason: 'rebase-conflict' };
				setMirror(transform, mapFrom, transform.steps.length - 1);
				committed.push(mapped);
			}
		} catch (cause) {
			return {
				status: 'rejected',
				reason: cause instanceof Error ? `invalid-step: ${cause.message}` : 'invalid-step',
			};
		}

		const committedBatch = freezeBatch({
			protocol: 1,
			sessionId: this.sessionId,
			batchId: batch.batchId,
			version: this.version,
			clientId: batch.clientId,
			steps: committed.map((step) => step.toJSON()),
		});
		const committedStart = transform.steps.length - committed.length;
		for (const [index, step] of committed.entries()) {
			this.records.push({ version: this.version, clientId: batch.clientId, step });
			this.version++;
			const doc =
				index === committed.length - 1
					? transform.doc
					: expectDefined(transform.docs[committedStart + index + 1], 'committed document');
			this.documents.set(this.version, doc);
		}
		this.requests.remember(batch.clientId, batch.batchId, fingerprint, committedBatch);
		this.pruneHistory();
		return { status: 'accepted', batch: committedBatch };
	}

	private pruneHistory() {
		while (this.records.length > this.stepHistoryLimit) {
			const removed = this.records.shift()!;
			for (const version of this.documents.keys())
				if (version <= removed.version) this.documents.delete(version);
		}
	}
}
