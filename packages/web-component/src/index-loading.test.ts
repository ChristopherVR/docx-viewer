// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { loadDocument } from '@christophervr/docx-document';
import { DocxEditorElement } from './index';
import { at, paragraphAt } from './test-support';

vi.mock('@christophervr/docx-document', () => ({ loadDocument: vi.fn() }));
const loadMock = vi.mocked(loadDocument);

describe('DocxEditorElement loading', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.resetAllMocks();
	});

	it('ignores stale loads and preserves the active imported session for save', async () => {
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		document.body.append(editor);
		let resolveOld!: (value: Awaited<ReturnType<typeof loadDocument>>) => void;
		const oldLoad = new Promise<Awaited<ReturnType<typeof loadDocument>>>((resolve) => {
			resolveOld = resolve;
		});
		const current = createDocument();
		at(paragraphAt(current.blocks, 0).runs, 0).text = 'current';
		const bytes = new Uint8Array([9, 8, 7]);
		const save = vi.fn().mockResolvedValue(bytes);
		loadMock.mockReturnValueOnce(oldLoad).mockResolvedValueOnce({ model: current, save });
		const first = editor.load(new Uint8Array([1]));
		await editor.load(new Uint8Array([2]));
		const stale = createDocument();
		at(paragraphAt(stale.blocks, 0).runs, 0).text = 'stale';
		resolveOld({ model: stale, save: vi.fn() });
		await first;
		expect(editor.documentModel?.blocks[0]).toMatchObject({ runs: [{ text: 'current' }] });
		await expect(editor.saveBytes()).resolves.toBe(bytes);
		expect(save).toHaveBeenCalledWith(current);
	});

	it('exposes import errors as events while rejecting the load promise', async () => {
		const editor = new DocxEditorElement();
		const error = new Error('bad package');
		loadMock.mockRejectedValueOnce(error);
		const handler = vi.fn();
		editor.addEventListener('document-error', handler);
		await expect(editor.load(new Uint8Array())).rejects.toBeTruthy();
		expect(handler).toHaveBeenCalledTimes(1);
		expect(at(at(handler.mock.calls, 0), 0).detail).toBeInstanceOf(Error);
		expect(at(at(handler.mock.calls, 0), 0).detail).toBe(error);
	});
});
