export { DocxEditorElement, registerDocxEditor } from './component';
export { DOCX_EDITOR_EVENTS } from './events';
export type { RibbonAction } from './ribbon';
export type { FileCommand, FileCommandDetail } from './file-commands';
export type { DocxEditorEventMap, DocxEditorEventName, DocxEditorEventDetail } from './events';
export { DOCX_EDITOR_ATTRIBUTES } from './editor-attributes';
export type { DocxEditorAttribute } from './editor-attributes';
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
export type { ReviewDisplayMode } from './review-display';
export type { RevisionRange } from './review-commands';
export { lightTheme, darkTheme, themeToCssVars } from './theme';
export type { EditorTheme, EditorThemeMode } from './theme';
