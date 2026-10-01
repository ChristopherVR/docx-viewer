// Surface every published framework package re-exports next to its component, so an application
// needs no second package to configure the editor or open a DOCX / legacy .doc file.
export * from '@christophervr/docx-web-component';
export { detectDocumentFormat, loadDocument } from '@christophervr/docx-document';
export type { DocumentFormat } from '@christophervr/docx-document';
export type {
	EditorBinding,
	EditorEventOptions,
	EditorHandle,
	EditorOptions,
	EditorProps,
} from './index';
