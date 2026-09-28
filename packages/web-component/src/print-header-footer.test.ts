// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	createDocument,
	loadDocx,
	type DocumentModel,
	type Paragraph,
	type SectionProperties,
} from '@christophervr/docx-core';
import type { LayoutPageBox } from '@christophervr/docx-layout';
import { decoratePages, headerFooterForPage, pageNumbers } from './print-header-footer';
import { at, must, paragraphAt } from './test-support';

const page = (index: number, sectionIndex: number, pageInSection: number): LayoutPageBox => ({
	index,
	sectionIndex,
	pageInSection,
	widthPx: 816,
	heightPx: 1056,
	marginTopPx: 96,
	marginRightPx: 96,
	marginBottomPx: 96,
	marginLeftPx: 96,
	columns: [],
});
const content = (text: string, field?: string) => ({
	blocks: [
		{
			type: 'paragraph' as const,
			id: `h-${text}`,
			runs: [{ text: 'Page ' }, { text, ...(field ? { field: { instr: field } } : {}) }],
		},
	],
});
function section(overrides: Partial<SectionProperties>): SectionProperties {
	return {
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
		...overrides,
	};
}

describe('Print Layout headers, footers and page fields', () => {
	it('numbers pages across sections, honoring restarts and number formats', () => {
		const model: DocumentModel = {
			...createDocument(),
			sections: [
				section({ pageNumbering: { format: 'lowerRoman' } }),
				section({ pageNumbering: { start: 1 } }),
			],
		};
		const pages = [page(0, 0, 0), page(1, 0, 1), page(2, 1, 0), page(3, 1, 1)];
		expect(pageNumbers(model, pages)).toEqual(['i', 'ii', '1', '2']);
	});

	it('chooses first, even and default headers like Word, inheriting from earlier sections', () => {
		const model: DocumentModel = {
			...createDocument(),
			evenAndOddHeaders: true,
			sections: [
				section({
					titlePage: true,
					headers: { first: content('first'), even: content('even'), default: content('odd') },
				}),
				section({}),
			],
		};
		const text = (p: LayoutPageBox, n: number) =>
			at(paragraphAt(must(headerFooterForPage(model, p, n, 'headers')).blocks, 0).runs, 1).text;
		expect(text(page(0, 0, 0), 1)).toBe('first');
		expect(text(page(1, 0, 1), 2)).toBe('even');
		expect(text(page(2, 0, 2), 3)).toBe('odd');
		expect(text(page(3, 1, 0), 4)).toBe('even');
	});

	it('fills PAGE and NUMPAGES per page when decorating sheets', () => {
		const model: DocumentModel = {
			...createDocument(),
			sections: [section({ footers: { default: content('1', 'PAGE') } })],
		};
		const footerBlocks = must(at(must(model.sections), 0).footers?.default).blocks;
		footerBlocks[0] = {
			type: 'paragraph',
			id: 'f1',
			runs: [
				{ text: '1', field: { instr: 'PAGE' } },
				{ text: ' of ' },
				{ text: '1', field: { instr: 'NUMPAGES \\* MERGEFORMAT' } },
			],
		};
		const pages = [page(0, 0, 0), page(1, 0, 1), page(2, 0, 2)];
		const sheets = pages.map(() => document.createElement('div'));
		decoratePages(model, pages, sheets);
		expect(sheets.map((sheet) => sheet.querySelector('.dve-print-footer')?.textContent)).toEqual([
			'1 of 3',
			'2 of 3',
			'3 of 3',
		]);
		expect(at(sheets, 0).querySelector<HTMLElement>('.dve-print-footer')!.style.bottom).toBe(
			'48px',
		);
	});

	it('tags complex field results with their field code when parsing', async () => {
		const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${w}"><w:body><w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const runs = (loaded.model.blocks[0] as Paragraph).runs.filter((run) => run.text);
		expect(runs).toEqual([
			{ text: 'Page ' },
			{ text: '7', field: { instr: 'PAGE' } },
			{ text: ' end' },
		]);
	});

	it('updates DATE fields in headers and footers to the layout time', () => {
		const model: DocumentModel = {
			...createDocument(),
			sections: [
				section({
					headers: {
						default: {
							blocks: [
								{
									type: 'paragraph',
									id: 'h1',
									runs: [{ text: '1/1/2020', field: { instr: 'DATE \\@ "yyyy-MM-dd"' } }],
								},
							],
						},
					},
				}),
			],
		};
		const sheets = [document.createElement('div')];
		decoratePages(model, [page(0, 0, 0)], sheets, new Date(2026, 8, 7));
		expect(at(sheets, 0).querySelector('.dve-print-header')?.textContent).toBe('2026-09-07');
	});
});

describe('Print Layout header and footer pictures', () => {
	const logo = {
		relId: 'rId1',
		partName: 'word/media/logo.png',
		contentType: 'image/png',
		widthPx: 80,
		heightPx: 40,
	};

	it('draws inline logos in the header and floating ones at their page position', () => {
		const model: DocumentModel = {
			...createDocument(),
			sections: [
				section({
					headers: {
						default: {
							blocks: [
								{
									type: 'paragraph',
									id: 'h1',
									runs: [
										{ text: '', image: logo },
										{ text: ' Company' },
										{
											text: '',
											image: {
												...logo,
												anchored: true,
												placement: {
													wrap: 'none',
													relativeFrom: 'page',
													align: 'right',
													relativeFromV: 'page',
													offsetYPx: 10,
												},
											},
										},
									],
								},
							],
						},
					},
				}),
			],
		};
		const sheets = [document.createElement('div')];
		decoratePages(model, [page(0, 0, 0)], sheets, new Date(), (partName) => `blob:${partName}`);
		const header = at(sheets, 0).querySelector('.dve-print-header')!;
		expect(header.textContent).toBe(' Company');
		const inline = header.querySelector('img')!;
		expect([inline.getAttribute('src'), inline.style.width]).toEqual([
			'blob:word/media/logo.png',
			'80px',
		]);
		const float = at(sheets, 0).querySelector<HTMLElement>(':scope > .dve-print-float')!;
		expect([float.style.left, float.style.top]).toEqual(['736px', '10px']);
	});
});
