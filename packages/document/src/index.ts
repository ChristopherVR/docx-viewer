import { loadDocx, type LoadedDocument } from '@christophervr/docx-core';
import { loadLegacyDoc } from '@christophervr/docx-legacy';
export type DocumentFormat = 'docx' | 'doc';
export function detectDocumentFormat(input: Uint8Array | ArrayBuffer): DocumentFormat {
	const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
	if (
		[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, index) => bytes[index] === value)
	)
		return 'doc';
	if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04)
		return 'docx';
	throw new Error('Unsupported file. Open a DOCX or Word 97-2003 DOC document.');
}
/** Sniff content, never trust a user-supplied filename extension. */
export async function loadDocument(input: Uint8Array | ArrayBuffer): Promise<LoadedDocument> {
	return detectDocumentFormat(input) === 'doc' ? loadLegacyDoc(input) : loadDocx(input);
}
