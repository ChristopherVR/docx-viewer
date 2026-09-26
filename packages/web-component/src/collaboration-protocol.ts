import type { Schema } from 'prosemirror-model';
import { Step } from 'prosemirror-transform';

export interface CollaborationConfig {
	sessionId: string;
	clientId: string;
	version?: number;
}

/** Version is the authority version immediately before applying these steps. */
export interface StepBatch {
	protocol: 1;
	sessionId: string;
	batchId: string;
	version: number;
	clientId: string;
	steps: readonly unknown[];
}

export type ParsedBatch = { batch: StepBatch; steps: Step[]; fingerprint: string };

export function validId(value: unknown, max = 160): value is string {
	return typeof value === 'string' && value.length > 0 && value.length <= max;
}

export function freezeBatch(batch: StepBatch): StepBatch {
	const freeze = (value: unknown): unknown => {
		if (value && typeof value === 'object' && !Object.isFrozen(value)) {
			for (const child of Object.values(value)) freeze(child);
			Object.freeze(value);
		}
		return value;
	};
	return freeze(batch) as StepBatch;
}

export function parseBatch(value: unknown, schema: Schema): ParsedBatch | string {
	if (!value || typeof value !== 'object' || Array.isArray(value)) return 'Batch must be an object';
	const input = value as Record<string, unknown>;
	if (input.protocol !== 1) return 'Unsupported collaboration protocol';
	if (!validId(input.sessionId) || !validId(input.batchId) || !validId(input.clientId))
		return 'Session, batch, and client IDs must be non-empty strings';
	if (!Number.isSafeInteger(input.version) || Number(input.version) < 0)
		return 'Version must be a non-negative safe integer';
	if (!Array.isArray(input.steps) || input.steps.length < 1 || input.steps.length > 500)
		return 'Batch must contain between 1 and 500 steps';
	let serialized: string;
	try {
		serialized = JSON.stringify(input.steps);
	} catch {
		return 'Steps are not serializable';
	}
	if (serialized.length > 1_000_000) return 'Step batch exceeds the 1 MB limit';
	try {
		const steps = input.steps.map((step) => Step.fromJSON(schema, step));
		const batch: StepBatch = {
			protocol: 1,
			sessionId: input.sessionId,
			batchId: input.batchId,
			version: Number(input.version),
			clientId: input.clientId,
			steps: input.steps,
		};
		return { batch, steps, fingerprint: JSON.stringify(batch) };
	} catch (cause) {
		return cause instanceof Error ? `Invalid step JSON: ${cause.message}` : 'Invalid step JSON';
	}
}
