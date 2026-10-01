import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import * as vanilla from './index';

const legacyFixture = new URL(
	'../../legacy/src/__tests__/fixtures/ole-word-97.doc',
	import.meta.url,
);

describe('@christophervr/docx-vanilla-viewer entry', () => {
	it('owns the web-component entry, the mount helper and the document loaders', () => {
		for (const name of [
			'mountEditor',
			'registerDocxEditor',
			'DocxEditorElement',
			'normalizeEditorLocale',
			'loadDocument',
			'detectDocumentFormat',
			'createCollaborationAuthority',
		] as const)
			expect(vanilla[name], name).toBeTypeOf('function');
	});

	it('imports without a DOM so server rendering can load the module', () => {
		expect(typeof globalThis.document).toBe('undefined');
	});

	it('opens a legacy .doc through the inlined legacy reader and keeps a no-op save byte-identical', async () => {
		const bytes = new Uint8Array(await readFile(legacyFixture));
		expect(vanilla.detectDocumentFormat(bytes)).toBe('doc');
		const loaded = await vanilla.loadDocument(bytes);
		expect(loaded.model.blocks.length).toBeGreaterThan(0);
		expect(await loaded.save()).toEqual(bytes);
	});
});
