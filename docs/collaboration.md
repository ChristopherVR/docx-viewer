# Collaboration protocol

The web component provides a transport-neutral collaborative editing foundation built on `prosemirror-collab`. It exchanges validated ProseMirror steps through an authority that assigns a single ordered version stream. It does not connect to a network or provide authentication, persistence, presence, or availability guarantees.

Try the [two-peer local collaboration demo](/demo/collaboration.html). The web component exposes `startCollaboration`, `getPendingCollaboration`, `receiveCollaboration`, and `stopCollaboration`. Hosts can use `createCollaborationAuthority(model, options)` to create the in-memory reference authority with the same schema and model conversion as the editor.

## Client and host flow

All participants editing the same document use the same `sessionId`; every editor incarnation uses a unique `clientId`. Reusing a client ID can make peers mistake each other's steps for acknowledgements. All participants must use the same schema and starting document. To join an existing authority, initialize the client with its current version and canonical document:

```ts
const client = new CollaborationClient({ sessionId, clientId, version: currentVersion });
const state = EditorState.create({ doc: canonicalDoc, plugins: [history(), client.plugin] });
```

For the custom element, call `editor.startCollaboration({ sessionId, clientId, version })` after loading the matching model snapshot. Keep each receiver mounted while it participates: `receiveCollaboration` applies through the live editor view. Listen for `collaboration-send`; submit each `StepBatch` to the authority and deliver its accepted batch to every editor with `receiveCollaboration(batch)`. Rejections such as `stale`, `out-of-order`, or `wrong-session` should trigger host-level recovery rather than being silently dropped.

After local transactions, call `client.createPendingBatch(state)`. It returns `null` when there are no unconfirmed steps, and otherwise returns an immutable envelope. Keep retrying that exact envelope until the authority accepts it; do not start another request while it is in flight. The method returns the same envelope until its matching accepted batch is received.

The host submits the envelope to its authority and relays the accepted `batch` to every participant, including the sender. The `version` on an accepted batch is the authority version immediately before the batch's steps. Participants pass the envelope to `client.receive(state, batch)`. On `applied`, dispatch the returned transaction through the editor's usual state-update path. This lets ordinary model updates, status events, and save behavior run for remote changes too. Applying remote batches does not consult the view's read-only setting; that setting blocks local edits only.

The protocol accepts at most 500 steps and 1,000,000 serialized JSON characters per batch. The reference `CollaborationAuthority` is an in-memory ordering and rebase implementation. It accepts current-version requests and maps stale requests through intervening accepted steps. Its default history retains 2,000 steps, and it caches 5,000 request acknowledgements for retries. A stale request older than retained history is rejected; a future version is rejected as out of order. The client remembers the most recent 1,000 received batch IDs to ignore duplicates. Hosts should provide catch-up or reload from their canonical snapshot after a gap, expired version, or rejected oversized batch.

## Presence

`PresenceClient` provides a separate transient channel for peer names, colors, cursors, and selections. It does not use or change the document step stream. Add its `.plugin` beside the collaboration plugin. `publish(state, { name, color })` returns a message for the host transport, or `null` while local document steps are still awaiting acknowledgement. `leave(state)` returns a removal message. After receiving a message, dispatch the transaction from `receive(state, message)` when its status is `applied`; ignore `duplicate` and recover from `stale`, `out-of-order`, or `wrong-session` according to the host policy.

Presence messages carry the shared session ID, sender client ID, committed collaboration version, and a monotonically increasing sequence per sender. Relay messages in order and republish after edit acknowledgement; positions are defined against the message's committed version. The plugin maps existing cursors and selections through subsequent document transactions, including edits received while the editor is read-only. It accepts names up to 80 printable characters, colors from `PRESENCE_PALETTE`, and at most 100 visible peers. Hosts should rate-limit and expire disconnected peers in their transport layer. Presence is not persisted, authenticated, or included in DOCX export.

## Stable identities

For the shared custom element, presence is enabled by `startCollaboration`:

```ts
import { PRESENCE_PALETTE } from 'docx-vanilla-viewer';

editor.addEventListener('presence-send', (event) => {
	transport.sendPresence(event.detail); // Application-owned transport
});
editor.publishPresence({ name: 'Ada', color: PRESENCE_PALETTE[0] });
// Incoming messages: editor.receivePresence(message)
// Explicit departure: editor.leavePresence()
```

The element republishes its stored profile after selection changes and edit
acknowledgements. It returns `out-of-order` for incoming selection positions while
local edits are pending; the host can request fresh presence after synchronization.
Leave messages contain no positions and can remove a peer at a different document
version, while still enforcing session and sender-sequence checks.

Paragraph and table IDs are document content and travel in ProseMirror steps. For local transactions, use `repairCollaborativeDocumentIds(state, clientId, idGenerator)` instead of the non-collaborative ID repair helper. Create `idGenerator` once with `createCollaborationIdGenerator(clientId)` and reuse it for that editor's lifetime; this prevents ID reuse after deletion and namespaces locally inserted table, cell, and paragraph IDs. The host must assign a distinct client ID to every live editor incarnation; the session ID and client ID are protocol identifiers, not credentials.

`stopCollaboration()` requires the pending edits to be acknowledged. If the host intentionally abandons them, call `stopCollaboration(true)` to explicitly discard the local transport queue. Disconnection or unmounting should not be treated as acknowledgement; retain the canonical snapshot/version and resume or recover through the host's transport policy.

## Authority responsibilities and limits

The application host owns transport, authentication and authorization, durable storage, snapshot distribution, client identity, reconnect policy, and export timing. The in-memory authority is useful for tests and a local demo, not as a multi-process or production server. DOCX export should use a canonical, sufficiently synchronized model if the application needs shared-save semantics; this library does not coordinate file locks or simultaneous exports. Remote peers receive only schema-supported ProseMirror steps, and the authority validates those steps before accepting them. Structural table commands are disabled during collaboration because the current table editor replaces table structure as a whole; granular cell-level transactions are needed before enabling those commands safely.
