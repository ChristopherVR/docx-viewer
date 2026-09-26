// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { history } from 'prosemirror-history';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { CollaborationAuthority, CollaborationClient, type StepBatch } from './collaboration';
import { schema } from './schema';

function documentWithText(text = 'abcdef') {
	return schema.nodes.doc.create(null, [
		schema.nodes.paragraph.create({ id: 'p1' }, text ? [schema.text(text)] : []),
	]);
}

function editorState(doc: ReturnType<typeof documentWithText>, client: CollaborationClient) {
	return EditorState.create({
		doc,
		selection: TextSelection.create(doc, 1),
		plugins: [history(), client.plugin],
	});
}

function local(state: EditorState, transaction: Transaction) {
	return state.apply(transaction);
}

function apply(client: CollaborationClient, state: EditorState, batch: StepBatch): EditorState {
	const result = client.receive(state, batch);
	expect(result.status).toBe('applied');
	if (result.status !== 'applied') throw new Error(`Could not apply batch: ${result.status}`);
	return state.apply(result.transaction);
}

function accept(authority: CollaborationAuthority, batch: StepBatch): StepBatch {
	const result = authority.submit(batch);
	if (result.status !== 'accepted')
		throw new Error(`Authority rejected batch: ${JSON.stringify(result)}`);
	return result.batch;
}

describe('ProseMirror collaboration protocol', () => {
	afterEach(() => document.body.replaceChildren());

	it('keeps generated batch IDs valid for a maximum-length client ID', () => {
		const doc = documentWithText();
		const clientId = 'x'.repeat(160);
		const client = new CollaborationClient({ sessionId: 'doc-long-client', clientId });
		const state = editorState(doc, client);
		const changed = local(state, state.tr.insertText('!', 2));
		const batch = client.createPendingBatch(changed)!;
		expect(batch.batchId.length).toBeLessThanOrEqual(160);
		expect(
			new CollaborationAuthority({ sessionId: 'doc-long-client', doc }).submit(batch).status,
		).toBe('accepted');
	});

	it('rebases concurrent inserts from two peers and acknowledges each batch', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'doc-a', doc });
		const a = new CollaborationClient({ sessionId: 'doc-a', clientId: 'peer-a' });
		const b = new CollaborationClient({ sessionId: 'doc-a', clientId: 'peer-b' });
		let stateA = editorState(doc, a);
		let stateB = editorState(doc, b);
		stateA = local(stateA, stateA.tr.insertText('A', 2));
		stateB = local(stateB, stateB.tr.insertText('B', 2));
		const requestA = a.createPendingBatch(stateA)!;
		const requestB = b.createPendingBatch(stateB)!;
		const committedA = accept(authority, requestA);
		stateA = apply(a, stateA, committedA);
		stateB = apply(b, stateB, committedA);
		const committedB = accept(authority, requestB);
		stateA = apply(a, stateA, committedB);
		stateB = apply(b, stateB, committedB);

		expect(authority.currentVersion).toBe(2);
		expect(authority.doc.textContent).toContain('A');
		expect(authority.doc.textContent).toContain('B');
		expect(stateA.doc.eq(stateB.doc)).toBe(true);
		expect(stateA.doc.eq(authority.doc)).toBe(true);
		expect(a.hasPendingAck).toBe(false);
		expect(b.hasPendingAck).toBe(false);
	});

	it('rebases concurrent deletion, insertion, and formatting steps', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'doc-b', doc });
		const a = new CollaborationClient({ sessionId: 'doc-b', clientId: 'left' });
		const b = new CollaborationClient({ sessionId: 'doc-b', clientId: 'right' });
		let stateA = editorState(doc, a);
		let stateB = editorState(doc, b);
		stateA = local(stateA, stateA.tr.delete(2, 3));
		stateB = local(stateB, stateB.tr.insertText('X', 5).addMark(1, 2, schema.marks.bold.create()));
		const requestA = a.createPendingBatch(stateA)!;
		const requestB = b.createPendingBatch(stateB)!;
		const committedA = accept(authority, requestA);
		stateA = apply(a, stateA, committedA);
		stateB = apply(b, stateB, committedA);
		const committedB = accept(authority, requestB);
		stateA = apply(a, stateA, committedB);
		stateB = apply(b, stateB, committedB);

		expect(stateA.doc.eq(authority.doc)).toBe(true);
		expect(stateB.doc.eq(authority.doc)).toBe(true);
		expect(authority.doc.textContent).toBe('acdXef');
		expect(authority.doc.firstChild?.child(0).marks.some((mark) => mark.type.name === 'bold')).toBe(
			true,
		);
	});

	it('deduplicates retransmitted submissions and reports stale and future versions', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'doc-c', doc });
		const client = new CollaborationClient({ sessionId: 'doc-c', clientId: 'peer' });
		let state = editorState(doc, client);
		state = local(state, state.tr.insertText('!', 2));
		const request = client.createPendingBatch(state)!;
		const accepted = authority.submit(request);
		expect(accepted.status).toBe('accepted');
		if (accepted.status !== 'accepted') throw new Error('Expected accepted batch');
		const duplicate = authority.submit(request);
		expect(duplicate).toEqual({ status: 'duplicate', batch: accepted.batch });
		expect(authority.submit({ ...request, steps: [] }).status).toBe('rejected');
		expect(authority.currentVersion).toBe(1);
		const old = { ...request, batchId: 'late', version: 0 };
		expect(authority.submit(old).status).toBe('accepted');
		expect(authority.submit({ ...request, batchId: 'future', version: 99 }).status).toBe(
			'rejected',
		);
		expect(authority.submit({ ...request, batchId: 'wrong', sessionId: 'other' })).toMatchObject({
			status: 'rejected',
			reason: 'wrong-session',
		});
	});

	it('handles duplicate, stale, out-of-order, and wrong-session inbound batches', () => {
		const doc = documentWithText();
		const client = new CollaborationClient({ sessionId: 'room', clientId: 'local' });
		let state = editorState(doc, client);
		const valid = {
			protocol: 1 as const,
			sessionId: 'room',
			batchId: 'server:1',
			version: 0,
			clientId: 'remote',
			steps: [state.tr.insertText('R', 2).steps[0].toJSON()],
		};
		const first = client.receive(state, valid);
		expect(first.status).toBe('applied');
		if (first.status === 'applied') state = state.apply(first.transaction);
		expect(client.receive(state, valid).status).toBe('duplicate');
		expect(client.receive(state, { ...valid, batchId: 'old:1' }).status).toBe('stale');
		expect(client.receive(state, { ...valid, batchId: 'future:1', version: 4 }).status).toBe(
			'out-of-order',
		);
		expect(client.receive(state, { ...valid, sessionId: 'wrong' }).status).toBe('wrong-session');
	});

	it('keeps one pending batch stable through retries and clears it on matching acknowledgement', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'ack', doc });
		const client = new CollaborationClient({ sessionId: 'ack', clientId: 'writer' });
		let state = editorState(doc, client);
		state = local(state, state.tr.insertText('!', 2));
		const pending = client.createPendingBatch(state)!;
		expect(client.hasPendingAck).toBe(true);
		expect(client.pendingStepCount(state)).toBe(1);
		expect(client.createPendingBatch(state)).toBe(pending);
		expect(Object.isFrozen(pending)).toBe(true);
		expect(Object.isFrozen(pending.steps)).toBe(true);
		const committed = accept(authority, pending);
		state = apply(client, state, committed);
		expect(client.hasPendingAck).toBe(false);
		expect(client.pendingStepCount(state)).toBe(0);
	});

	it('acknowledges only the matching prefix when local edits continue in flight', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'in-flight', doc });
		const client = new CollaborationClient({ sessionId: 'in-flight', clientId: 'writer' });
		let state = editorState(doc, client);
		state = local(state, state.tr.insertText('A', 2));
		const first = client.createPendingBatch(state)!;
		state = local(state, state.tr.insertText('B', 3));
		const committed = accept(authority, first);
		state = apply(client, state, committed);
		expect(client.hasPendingAck).toBe(false);
		expect(client.pendingStepCount(state)).toBeGreaterThan(0);
		const second = client.createPendingBatch(state)!;
		expect(second.batchId).not.toBe(first.batchId);
		expect(second.steps).toHaveLength(1);
	});

	it('does not clear local steps for a mismatched self acknowledgement', () => {
		const doc = documentWithText();
		const client = new CollaborationClient({ sessionId: 'forged-ack', clientId: 'writer' });
		let state = editorState(doc, client);
		state = local(state, state.tr.insertText('A', 2));
		const pending = client.createPendingBatch(state)!;
		const wrong = {
			...pending,
			version: 0,
			steps: [state.tr.insertText('X', 3).steps[0].toJSON()],
		};
		expect(client.receive(state, wrong)).toMatchObject({ status: 'invalid' });
		expect(client.hasPendingAck).toBe(true);
		expect(client.pendingStepCount(state)).toBe(1);
	});

	it('applies remote commits to a readonly view', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'readonly', doc });
		const remote = new CollaborationClient({ sessionId: 'readonly', clientId: 'remote' });
		const readonly = new CollaborationClient({ sessionId: 'readonly', clientId: 'viewer' });
		let remoteState = editorState(doc, remote);
		remoteState = local(remoteState, remoteState.tr.insertText('R', 2));
		let readonlyState = editorState(doc, readonly);
		const host = document.createElement('div');
		document.body.append(host);
		const view = new EditorView(host, {
			state: readonlyState,
			editable: () => false,
		});
		const committed = accept(authority, remote.createPendingBatch(remoteState)!);
		const result = readonly.receive(readonlyState, committed);
		expect(result.status).toBe('applied');
		if (result.status === 'applied') {
			readonlyState = readonlyState.apply(result.transaction);
			view.updateState(readonlyState);
		}
		expect(view.editable).toBe(false);
		expect(view.state.doc.textContent).toContain('R');
		view.destroy();
	});

	it('rejects malformed step JSON before changing authority or client state', () => {
		const doc = documentWithText();
		const authority = new CollaborationAuthority({ sessionId: 'validate', doc });
		const malformed = {
			protocol: 1,
			sessionId: 'validate',
			batchId: 'bad',
			version: 0,
			clientId: 'peer',
			steps: [{ stepType: 'unknown' }],
		};
		expect(authority.submit(malformed).status).toBe('rejected');
		expect(authority.currentVersion).toBe(0);
		const invalidPosition = {
			...malformed,
			batchId: 'negative-position',
			steps: [{ stepType: 'replace', from: -1, to: 0 }],
		};
		expect(authority.submit(invalidPosition).status).toBe('rejected');
		expect(authority.currentVersion).toBe(0);
		const client = new CollaborationClient({ sessionId: 'validate', clientId: 'local' });
		const state = editorState(doc, client);
		expect(client.receive(state, malformed).status).toBe('invalid');
		expect(client.receive(state, { ...invalidPosition, clientId: 'remote' }).status).toBe(
			'invalid',
		);
		expect(state.doc.eq(doc)).toBe(true);
	});
});
