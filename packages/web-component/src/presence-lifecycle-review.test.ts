// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { collab, receiveTransaction } from 'prosemirror-collab';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorPresence } from './editor-presence';
import { PresenceClient, PRESENCE_PALETTE } from './presence';
import { schema } from './schema';

function documentWithText(text = 'abcd') {
	return schema.node('doc', null, [schema.node('paragraph', { id: 'p1' }, [schema.text(text)])]);
}

describe('presence lifecycle regression review', () => {
	it('does not retain a rejected profile and reschedule it as an uncaught microtask error', () => {
		const doc = documentWithText();
		const state = EditorState.create({
			doc,
			selection: TextSelection.create(doc, 1),
			plugins: [collab()],
		});
		const presence = new EditorPresence(
			{ sessionId: 'review-doc', clientId: 'local' },
			document.createElement('div'),
			() => ({ state }) as never,
		);
		let scheduled: (() => void) | undefined;
		vi.stubGlobal('queueMicrotask', (callback: () => void) => (scheduled = callback));
		try {
			expect(() => presence.publish({ name: ' ', color: '#fff' })).toThrow('Presence name');
			presence.schedule();
			expect(scheduled).toBeUndefined();
		} finally {
			vi.unstubAllGlobals();
		}
	});

	it('accepts a fresh leave message after document version advances', () => {
		const doc = documentWithText();
		const senderState = EditorState.create({ doc, plugins: [collab({ clientID: 'alice' })] });
		const sender = new PresenceClient({ sessionId: 'review-doc', clientId: 'alice' });
		const receiver = new PresenceClient({ sessionId: 'review-doc', clientId: 'bob' });
		let receiverState = EditorState.create({
			doc,
			plugins: [collab({ clientID: 'bob' }), receiver.plugin],
		});
		const active = sender.publish(senderState, { name: 'Alice', color: PRESENCE_PALETTE[0] })!;
		const joined = receiver.receive(receiverState, active);
		expect(joined.status).toBe('applied');
		if (joined.status !== 'applied') return;
		receiverState = receiverState.apply(joined.transaction);
		const remoteStep = senderState.tr.insertText('x', 2).steps;
		receiverState = receiverState.apply(receiveTransaction(receiverState, remoteStep, ['alice']));
		const left = sender.leave(senderState);
		left.sequence = active.sequence + 1;
		expect(receiver.receive(receiverState, left).status).toBe('applied');
		const newerSelection = {
			...active,
			sequence: left.sequence + 1,
			version: 1,
			anchor: 2,
			head: 2,
		};
		expect(receiver.receive(receiverState, newerSelection).status).toBe('applied');
		expect(receiver.receive(receiverState, left).status).toBe('stale');
	});
});
