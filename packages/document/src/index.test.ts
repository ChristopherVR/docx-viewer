import { describe, expect, it } from 'vitest';
import { createDocument, saveDocx } from '@christophervr/docx-core';
import { detectDocumentFormat, loadDocument } from './index';
describe('document routing', () => {
	it('routes DOCX by content and respects typed-array slices', async () => {
		const bytes = await saveDocx(createDocument());
		const larger = new Uint8Array(bytes.length + 16);
		larger.set(bytes, 8);
		const view = larger.subarray(8, 8 + bytes.length);
		expect(detectDocumentFormat(view)).toBe('docx');
		expect((await loadDocument(view)).model.blocks).toHaveLength(1);
	});
	it('rejects unknown and truncated inputs with a useful error', async () => {
		await expect(loadDocument(new Uint8Array([0x50, 0x4b]))).rejects.toThrow('Unsupported file');
		await expect(loadDocument(new Uint8Array([0x50, 0x4b, 3, 4]))).rejects.toThrow();
	});
});
