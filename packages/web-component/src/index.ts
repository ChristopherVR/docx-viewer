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
export { PresenceClient, PRESENCE_PALETTE } from './presence';
export type { PresenceMessage, PresenceConfig, PresenceReceiveResult } from './presence';
export { normalizeEditorLocale } from './localization';
export type { EditorLocale } from './localization';
export {
	createCollaborationIdGenerator,
	repairCollaborativeDocumentIds,
} from './collaboration-identity';
