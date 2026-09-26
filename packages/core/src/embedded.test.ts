import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { readOleDocumentParagraphs, writeOleDocumentParagraphEdit } from './embedded.js';

async function docxFixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>First</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('embedded DOCX compatibility API', () => {
	it('reads body paragraphs and writes edits through the shared DOCX codec', async () => {
		const source = await docxFixture();
		expect(await readOleDocumentParagraphs(source)).toEqual(['First', 'Second']);
		const edited = await writeOleDocumentParagraphEdit(source, 1, 'Edited second');
		expect(await readOleDocumentParagraphs(edited)).toEqual(['First', 'Edited second']);
		expect(await readOleDocumentParagraphs(source)).toEqual(['First', 'Second']);
	});
});
