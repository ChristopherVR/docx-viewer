import { Decoration, DecorationSet } from 'prosemirror-view';
import type { EditorState } from 'prosemirror-state';
import { translateTemplate, type EditorLocale } from './localization';
import type { PeerPresence } from './presence';

/** Inline selection highlights and labelled cursors for remote peers. */
export function peerDecorations(
	doc: EditorState['doc'],
	peers: Map<string, PeerPresence>,
	locale: EditorLocale,
): DecorationSet {
	const decorations: Decoration[] = [];
	for (const peer of peers.values()) {
		const from = Math.min(peer.anchor, peer.head);
		const to = Math.max(peer.anchor, peer.head);
		if (from < to) {
			decorations.push(
				Decoration.inline(from, to, {
					class: 'dve-peer-selection',
					style: `background-color: ${peer.color}33; box-shadow: inset 0 -2px ${peer.color}`,
					'aria-label': translateTemplate(locale, 'presence.selection', { name: peer.name }),
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
		cursor.setAttribute(
			'aria-label',
			translateTemplate(locale, 'presence.cursor', { name: peer.name }),
		);
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
				key: `peer-${peer.clientId}-${peer.sequence}-${locale}`,
				side: 1,
			}),
		);
	}
	return DecorationSet.create(doc, decorations);
}
