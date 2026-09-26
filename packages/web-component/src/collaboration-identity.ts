import type { EditorState, Transaction } from 'prosemirror-state';

function clientKey(clientId: string): string {
	if (!clientId || clientId.length > 160) throw new Error('Invalid collaboration client ID');
	return Array.from(clientId, (character) => character.codePointAt(0)!.toString(16)).join('_');
}

/** Generates collision-resistant document IDs within one unique client session. */
export function createCollaborationIdGenerator(clientId: string): (kind: string) => string {
	const key = clientKey(clientId);
	let counter = 0;
	return (kind) => {
		const safeKind = kind.replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 24) || 'item';
		return `dve-${safeKind}-${key}-${++counter}`;
	};
}

/** Repairs missing or duplicate block IDs on local transactions. */
export function repairCollaborativeDocumentIds(
	state: EditorState,
	clientId: string,
	idGenerator = createCollaborationIdGenerator(clientId),
): Transaction | null {
	const reserved = new Set<string>();
	const seen = new Set<string>();
	state.doc.descendants((node) => {
		if ((node.type.name === 'paragraph' || node.type.name === 'table') && node.attrs.id)
			reserved.add(String(node.attrs.id));
	});
	let transaction = state.tr;
	state.doc.descendants((node, pos) => {
		if (node.type.name !== 'paragraph' && node.type.name !== 'table') return;
		const id = String(node.attrs.id || '');
		if (id && !seen.has(id)) {
			seen.add(id);
			return;
		}
		let generated = '';
		const kind = node.type.name === 'table' ? 'table' : 'paragraph';
		do generated = idGenerator(kind);
		while (reserved.has(generated));
		reserved.add(generated);
		seen.add(generated);
		transaction = transaction.setNodeMarkup(pos, undefined, { ...node.attrs, id: generated });
	});
	return transaction.docChanged ? transaction : null;
}

/** @deprecated Use repairCollaborativeDocumentIds to include table IDs as well. */
export const repairCollaborativeParagraphIds = repairCollaborativeDocumentIds;
