// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { collab } from 'prosemirror-collab';
import { EditorState, Plugin, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import {
	getPresenceDecorations,
	PresenceClient,
	PRESENCE_PALETTE,
	refreshPresenceLabels,
	type PresenceMessage,
} from './presence';
import { schema } from './schema';
import type { EditorLocale } from './localization';

function createDoc(text = 'abcdef') {
	return schema.nodes.doc.create(null, [
		schema.nodes.paragraph.create({ id: 'p' }, text ? [schema.text(text)] : []),
	]);
}

function createState(doc: ReturnType<typeof createDoc>, clientId: string, plugins: Plugin[] = []) {
	return EditorState.create({
		doc,
		selection: TextSelection.create(doc, 1),
		plugins: [collab({ clientID: clientId }), ...plugins],
	});
}

describe('transport-neutral collaboration presence', () => {
	it('maps peer selections through document transactions and removes peers on leave', () => {
		const doc = createDoc();
		const sender = new PresenceClient({ sessionId: 'doc', clientId: 'alice' });
		const receiver = new PresenceClient({ sessionId: 'doc', clientId: 'bob' });
		let state = createState(doc, 'bob', [receiver.plugin]);
		const message = sender.publish(
			EditorState.create({
				doc,
				selection: TextSelection.create(doc, 3, 5),
				plugins: [collab({ clientID: 'alice' })],
			}),
			{ name: 'Alice', color: PRESENCE_PALETTE[0] },
		);
		const received = receiver.receive(state, message);
		expect(received.status).toBe('applied');
		if (received.status !== 'applied') return;
		state = state.apply(received.transaction);
		const initial = getPresenceDecorations(state).find();
		expect(initial.map((decoration) => [decoration.from, decoration.to])).toContainEqual([3, 5]);
		const insert = state.tr.insertText('X', 2);
		state = state.apply(insert);
		const mapped = getPresenceDecorations(state).find();
		expect(mapped.map((decoration) => [decoration.from, decoration.to])).toContainEqual([4, 6]);
		const leave = sender.leave(createState(doc, 'alice'));
		const removed = receiver.receive(state, leave);
		expect(removed.status).toBe('applied');
		if (removed.status === 'applied') state = state.apply(removed.transaction);
		expect(getPresenceDecorations(state).find()).toHaveLength(0);
	});

	it('maps a collapsed cursor without expanding it and rejects older same-version presence', () => {
		const doc = createDoc();
		const sender = new PresenceClient({ sessionId: 'doc', clientId: 'alice' });
		const recipient = new PresenceClient({ sessionId: 'doc', clientId: 'bob' });
		let state = createState(doc, 'bob', [recipient.plugin]);
		const source = createState(doc, 'alice');
		const old = sender.publish(source, { name: 'Alice', color: PRESENCE_PALETTE[0] })!;
		const newest = sender.publish(
			source.apply(source.tr.setSelection(TextSelection.create(doc, 4))),
			{
				name: 'Alice',
				color: PRESENCE_PALETTE[0],
			},
		)!;
		const accepted = recipient.receive(state, newest);
		expect(accepted.status).toBe('applied');
		if (accepted.status === 'applied') state = state.apply(accepted.transaction);
		expect(recipient.receive(state, old).status).toBe('stale');
		const insertAtCursor = state.tr.insertText('X', 4);
		state = state.apply(insertAtCursor);
		const cursor = getPresenceDecorations(state)
			.find()
			.find((item) => String(item.spec.key).startsWith('peer-'));
		expect(cursor?.from).toBe(4);
	});

	it('validates session, version, range, name, and palette; supports view read-only', () => {
		const doc = createDoc();
		const recipient = new PresenceClient({ sessionId: 'doc', clientId: 'bob' });
		const sender = new PresenceClient({ sessionId: 'doc', clientId: 'alice' });
		const state = createState(doc, 'bob', [recipient.plugin]);
		const view = new EditorView(document.createElement('div'), {
			state,
			editable: () => false,
		});
		const valid = sender.publish(createState(doc, 'alice'), {
			name: 'Alice',
			color: PRESENCE_PALETTE[1],
		});
		expect(
			sender.publish(state.apply(state.tr.insertText('pending', 2)), {
				name: 'Alice',
				color: PRESENCE_PALETTE[1],
			}),
		).toBeNull();
		expect(recipient.receive(state, { ...valid, sessionId: 'other' }).status).toBe('wrong-session');
		expect(recipient.receive(state, { ...valid, version: 1 }).status).toBe('out-of-order');
		expect(recipient.receive(state, { ...valid, version: -1 }).status).toBe('invalid');
		expect(recipient.receive(state, { ...valid, head: doc.content.size + 1 }).status).toBe(
			'invalid',
		);
		expect(recipient.receive(state, { ...valid, name: '   ' }).status).toBe('invalid');
		expect(recipient.receive(state, { ...valid, color: '#ffffff' }).status).toBe('invalid');
		const applied = recipient.receive(view.state, valid);
		expect(applied.status).toBe('applied');
		if (applied.status === 'applied') view.dispatch(applied.transaction);
		expect(getPresenceDecorations(view.state).find().length).toBeGreaterThan(0);
		view.destroy();
	});

	it('ignores retransmissions and acknowledges the last peer state with leave', () => {
		const doc = createDoc();
		const recipient = new PresenceClient({ sessionId: 'doc', clientId: 'bob' });
		const sender = new PresenceClient({ sessionId: 'doc', clientId: 'alice' });
		let state = createState(doc, 'bob', [recipient.plugin]);
		const local = createState(doc, 'alice');
		const message = sender.publish(local, { name: 'Alice', color: PRESENCE_PALETTE[2] });
		const first = recipient.receive(state, message);
		expect(first.status).toBe('applied');
		if (first.status === 'applied') state = state.apply(first.transaction);
		expect(recipient.receive(state, message).status).toBe('duplicate');
		const leave: PresenceMessage = sender.leave(local);
		const removed = recipient.receive(state, leave);
		expect(removed.status).toBe('applied');
		if (removed.status === 'applied') state = state.apply(removed.transaction);
		expect(recipient.receive(state, leave).status).toBe('duplicate');
	});

	it('re-renders peer cursor labels as soon as the display locale changes', () => {
		const doc = createDoc();
		let locale: EditorLocale = 'en';
		const sender = new PresenceClient({ sessionId: 'doc', clientId: 'alice' });
		const recipient = new PresenceClient({
			sessionId: 'doc',
			clientId: 'bob',
			locale: () => locale,
		});
		const host = document.createElement('div');
		const view = new EditorView(host, {
			state: createState(doc, 'bob', [recipient.plugin]),
		});
		const message = sender.publish(createState(doc, 'alice'), {
			name: 'Alice',
			color: PRESENCE_PALETTE[0],
		});
		const received = recipient.receive(view.state, message);
		if (received.status === 'applied') view.dispatch(received.transaction);
		const label = () => host.querySelector('.dve-peer-cursor')?.getAttribute('aria-label');
		const english = label();
		expect(english).toContain('Alice');
		locale = 'de';
		expect(label()).toBe(english);
		view.dispatch(refreshPresenceLabels(view.state));
		expect(label()).not.toBe(english);
		expect(label()).toContain('Alice');
		view.destroy();
	});
});
