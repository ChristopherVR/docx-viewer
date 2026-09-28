// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from '@christophervr/docx-core';
import { DocxEditorElement, registerDocxEditor } from './index';
import { placementClass } from './inline-content-schema';
import { at } from './test-support';

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
		expect(at(runs, 1).link?.href).toBe('javascript:alert(1)');
	});
});

describe('pictures in headers', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.restoreAllMocks();
	});

	it('shows header pictures from package media in the header preview', () => {
		URL.createObjectURL = vi.fn(() => 'blob:header-logo');
		URL.revokeObjectURL = vi.fn();
		const model = createDocument();
		model.sections = [
			{
				endsAtBlockId: 'p1',
				type: 'nextPage',
				pageWidthTwips: 12240,
				pageHeightTwips: 15840,
				orientation: 'portrait',
				marginTopTwips: 1440,
				marginRightTwips: 1440,
				marginBottomTwips: 1440,
				marginLeftTwips: 1440,
				columns: { count: 1, equalWidth: true },
				headers: {
					default: {
						partName: 'word/header1.xml',
						blocks: [
							{
								type: 'paragraph',
								id: 'h1',
								runs: [
									{
										text: '',
										image: {
											relId: 'rId7',
											partName: 'word/media/logo.png',
											contentType: 'image/png',
											widthPx: 30,
											heightPx: 30,
										},
									},
								],
							},
						],
					},
				},
			},
		];
		const editor = document.createElement('docx-editor') as DocxEditorElement;
		editor.setLoadedDocument({
			model,
			media: new Map([['word/media/logo.png', new Uint8Array([137, 80, 78, 71])]]),
			save: async () => new Uint8Array(),
		});
		document.body.append(editor);
		const image = editor.shadowRoot!.querySelector<HTMLImageElement>(
			'.dve-header img[data-docx-image]',
		);
		expect(image?.getAttribute('src')).toBe('blob:header-logo');
	});
});

describe('floating picture classes', () => {
	it('approximates Word wrapping with floats and blocks', () => {
		const json = (value: object) => JSON.stringify(value);
		expect(placementClass(json({ wrap: 'square', align: 'right' }))).toBe('dve-float-right');
		expect(placementClass(json({ wrap: 'tight', offsetXPx: 20 }))).toBe('dve-float-left');
		expect(placementClass(json({ wrap: 'topAndBottom', align: 'center' }))).toBe(
			'dve-float-block dve-float-block-center',
		);
		expect(placementClass(json({ wrap: 'none', behindText: true }))).toBe('');
		expect(placementClass(null)).toBe('');
	});
});
