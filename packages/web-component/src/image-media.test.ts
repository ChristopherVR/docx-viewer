// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';

registerDocxEditor();

describe('inline picture rendering', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.restoreAllMocks();
	});

	it('renders pictures from package media and releases object URLs on disconnect', () => {
		const created: string[] = [];
		const revoked: string[] = [];
		URL.createObjectURL = vi.fn(() => {
			created.push(`blob:image-${created.length}`);
			return created.at(-1)!;
		});
		URL.revokeObjectURL = vi.fn((url: string) => void revoked.push(url));
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [
				{
					text: '',
					image: {
						relId: 'rId5',
						partName: 'word/media/image1.png',
						contentType: 'image/png',
						widthPx: 40,
						heightPx: 20,
						altText: 'Logo',
					},
				},
				{ text: ' caption' },
			],
		};
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.setLoadedDocument({
			model,
			media: new Map([['word/media/image1.png', new Uint8Array([137, 80, 78, 71])]]),
			save: async () => new Uint8Array(),
		});
		document.body.append(editor);
		const image = editor.shadowRoot!.querySelector<HTMLImageElement>('img[data-docx-image]')!;
		expect(image.getAttribute('src')).toBe('blob:image-0');
		expect(image.getAttribute('alt')).toBe('Logo');
		expect(image.getAttribute('width')).toBe('40');
		editor.remove();
		expect(revoked).toEqual(['blob:image-0']);
	});
});

describe('hyperlink rendering', () => {
	afterEach(() => document.body.replaceChildren());

	it('never renders a non-web link target as a navigable href', () => {
		const model = createDocument();
		model.blocks[0] = {
			type: 'paragraph',
			id: 'p1',
			runs: [
				{ text: 'safe', link: { href: 'https://example.com/' } },
				{ text: ' unsafe', link: { href: 'javascript:alert(1)' } },
			],
		};
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.documentModel = model;
		document.body.append(editor);
		const links = [...editor.shadowRoot!.querySelectorAll<HTMLAnchorElement>('a[data-docx-link]')];
		expect(links.map((link) => link.getAttribute('href'))).toEqual(['https://example.com/', '']);
		const runs = (editor.documentModel!.blocks[0] as { runs: { link?: { href?: string } }[] }).runs;
		expect(runs[1].link?.href).toBe('javascript:alert(1)');
	});
});
