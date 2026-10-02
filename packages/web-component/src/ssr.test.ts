// @vitest-environment node
import { expect, it, vi } from 'vitest';
import { createDocument } from 'docx-core';

vi.mock('@christophervr/ooxml-core/docx/load', () => ({
	loadDocument: async () => ({ model: createDocument(), save: async () => new Uint8Array() }),
}));

it('can be imported and registered during server rendering', async () => {
	const { registerDocxEditor } = await import('./index');
	expect(() => registerDocxEditor()).not.toThrow();
}, 30_000);
