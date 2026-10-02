// Surface every published framework package re-exports next to its component, so an application
// needs no second package to build a document model, configure the editor or open a DOCX / legacy
// .doc file (the model API comes from `@christophervr/docx-core`, which installs with the editor).
export * from '@christophervr/docx-core';
export * from '@christophervr/docx-web-component';
export { detectDocumentFormat, loadDocument } from '@christophervr/ooxml-core/docx/load';
export type { DocumentFormat } from '@christophervr/ooxml-core/docx/load';
export type {
	EditorBinding,
	EditorEventOptions,
	EditorHandle,
	EditorOptions,
	EditorProps,
} from './index';
