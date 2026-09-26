import './style.css';
import './collaboration.css';
import { createDocument } from '@christophervr/docx-core';
import { createCollaborationAuthority, type StepBatch } from '@christophervr/docx-web-component';
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
const handles = await Promise.all(
	['peer-a', 'peer-b'].map((id) =>
		mountFramework(document.getElementById(id)!, { documentModel: structuredClone(initial) }),
	),
);
const peers = handles.map((handle) => handle.element);
const queue = new Map<string, StepBatch>();
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
		queue.set(batch.batchId, batch);
		if (paused) renderStatus();
		else drain();
	});
	peer.startCollaboration({ sessionId, clientId: `editor-${index + 1}` });
});
pause.addEventListener('click', () => {
	paused = !paused;
	pause.textContent = paused ? 'Resume delivery' : 'Pause delivery';
	pause.setAttribute('aria-pressed', String(paused));
	if (paused) renderStatus();
	else drain();
});
document.getElementById('peer-readonly')!.addEventListener('change', (event) => {
	peers[1].readOnly = (event.target as HTMLInputElement).checked;
});
renderStatus();
