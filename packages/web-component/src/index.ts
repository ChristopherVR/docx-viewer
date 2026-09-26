export { DocxEditorElement, registerDocxEditor } from './component';
export { CollaborationClient, CollaborationAuthority } from './collaboration';
export type {
	CollaborationConfig,
	StepBatch,
	ClientReceiveResult,
	AuthorityResult,
} from './collaboration';
export type { CollaborationAuthorityConfig } from './collaboration';
export { createCollaborationAuthority } from './collaboration-model';
export {
	createCollaborationIdGenerator,
	repairCollaborativeDocumentIds,
} from './collaboration-identity';
