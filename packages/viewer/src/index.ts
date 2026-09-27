export { createDocument, loadDocx, saveDocx } from '@christophervr/docx-core';
export type {
	DocumentModel,
	LoadedDocument,
	Paragraph,
	TextRun,
	Table,
} from '@christophervr/docx-core';
export { detectDocumentFormat, loadDocument } from '@christophervr/docx-document';
export type { DocumentFormat } from '@christophervr/docx-document';
export { mountEditor } from '@christophervr/docx-bindings';
export type { EditorBinding, EditorHandle, EditorOptions } from '@christophervr/docx-bindings';
export { registerDocxEditor } from '@christophervr/docx-web-component';
export type { DocxEditorElement } from '@christophervr/docx-web-component';
