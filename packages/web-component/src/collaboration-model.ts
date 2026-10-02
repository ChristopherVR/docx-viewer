import type { DocumentModel } from 'docx-core';
import { CollaborationAuthority, type CollaborationAuthorityConfig } from './collaboration';
import { modelToDoc } from './model-adapter';

/** Build a reference authority using exactly the shared editor's document schema. */
export function createCollaborationAuthority(
	model: DocumentModel,
	options: Omit<CollaborationAuthorityConfig, 'doc'>,
): CollaborationAuthority {
	return new CollaborationAuthority({ ...options, doc: modelToDoc(model) });
}
