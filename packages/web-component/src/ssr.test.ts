// @vitest-environment node
import { expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';

vi.mock('@christophervr/docx-document', () => ({
	loadDocument: async () => ({ model: createDocument(), save: async () => new Uint8Array() }),
}));

it('can be imported and registered during server rendering', async () => {
	const { registerDocxEditor } = await import('./index');
	expect(() => registerDocxEditor()).not.toThrow();
}, 30_000);
