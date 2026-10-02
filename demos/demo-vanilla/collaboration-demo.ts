import './style.css';
import './collaboration.css';
import { createDocument } from 'docx-core';
import {
	createCollaborationAuthority,
	type StepBatch,
	type PresenceMessage,
} from 'docx-web-component';
import { mountFramework } from './framework';
import { initTheme } from './theme';

initTheme();
const initial = createDocument();
initial.page = {
	...initial.page,
	width: 560,
	marginLeft: 40,
	marginRight: 40,
	marginTop: 40,
	marginBottom: 40,
};
initial.blocks = [{ type: 'paragraph', id: 'shared-intro', runs: [{ text: 'Shared document' }] }];
const sessionId = 'local-coauthor-demo';
const authority = createCollaborationAuthority(initial, { sessionId });
const status = document.getElementById('collaboration-status')!;
const pause = document.getElementById('pause-delivery') as HTMLButtonElement;
const guestFramework = new URLSearchParams(location.search).get('guest') || undefined;
const handles = await Promise.all([
	mountFramework(document.getElementById('peer-a')!, { documentModel: structuredClone(initial) }),
	mountFramework(
		document.getElementById('peer-b')!,
		{ documentModel: structuredClone(initial) },
		guestFramework,
	),
]);
const peers = handles.map((handle) => handle.element);
const queue = new Map<string, StepBatch>();
const presenceQueue = new Map<string, PresenceMessage>();
const profiles = [
	{
		name: (document.getElementById('peer-a-name') as HTMLInputElement).value,
		color: (document.getElementById('peer-a-color') as HTMLSelectElement).value,
	},
	{
		name: (document.getElementById('peer-b-name') as HTMLInputElement).value,
		color: (document.getElementById('peer-b-color') as HTMLSelectElement).value,
	},
];
let paused = false;
let draining = false;

function renderStatus() {
	status.textContent = paused
		? `Delivery paused · ${queue.size} queued batches`
		: `Synced · version ${authority.currentVersion}`;
}

function drain() {
	if (paused || draining) return;
	draining = true;
	try {
		for (const [id, batch] of queue) {
			const result = authority.submit(batch);
			if (result.status === 'rejected')
				throw new Error(`Authority rejected an edit: ${result.reason}`);
			for (const peer of peers) {
				const received = peer.receiveCollaboration(result.batch);
				if (received !== 'applied' && received !== 'duplicate')
					throw new Error(`Peer could not receive edit: ${received}`);
			}
			queue.delete(id);
		}
		for (const [id, message] of presenceQueue) {
			let retry = false;
			for (const peer of peers) {
				const received = peer.receivePresence(message);
				if (received === 'out-of-order') {
					retry = true;
					continue;
				}
				if (received !== 'applied' && received !== 'duplicate' && received !== 'stale')
					throw new Error(`Peer could not receive presence: ${received}`);
			}
			if (!retry) presenceQueue.delete(id);
		}
		renderStatus();
	} catch (error) {
		status.textContent = error instanceof Error ? error.message : String(error);
	} finally {
		draining = false;
	}
}

peers.forEach((peer, index) => {
	peer.addEventListener('collaboration-send', (event) => {
		const batch = (event as CustomEvent<StepBatch>).detail;
		queue.set(`${batch.clientId}\u0000${batch.batchId}`, batch);
		if (paused) renderStatus();
		else drain();
	});
	peer.addEventListener('presence-send', (event) => {
		const message = (event as CustomEvent<PresenceMessage>).detail;
		presenceQueue.set(message.clientId, message);
		if (paused) renderStatus();
		else drain();
	});
	peer.startCollaboration({ sessionId, clientId: `editor-${index + 1}` });
	const name = document.getElementById(`peer-${index === 0 ? 'a' : 'b'}-name`)!;
	const color = document.getElementById(`peer-${index === 0 ? 'a' : 'b'}-color`)!;
	const updateProfile = () => {
		const nameValue = (name as HTMLInputElement).value.trim();
		if (!nameValue) return;
		profiles[index] = {
			name: nameValue,
			color: (color as HTMLSelectElement).value,
		};
		const profile = profiles[index];
		if (profile) peer.publishPresence(profile);
	};
	name.addEventListener('input', updateProfile);
	color.addEventListener('change', updateProfile);
});
peers.forEach((peer, index) => {
	const profile = profiles[index];
	if (profile) peer.publishPresence(profile);
});
pause.addEventListener('click', () => {
	paused = !paused;
	pause.textContent = paused ? 'Resume delivery' : 'Pause delivery';
	pause.setAttribute('aria-pressed', String(paused));
	if (paused) renderStatus();
	else drain();
});
document.getElementById('peer-readonly')!.addEventListener('change', (event) => {
	const second = peers[1];
	if (second) second.readOnly = (event.target as HTMLInputElement).checked;
});
renderStatus();
