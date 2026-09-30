import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocumentModel } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const load = async (sectPr: string) => {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>one</w:t></w:r></w:p>${sectPr}</w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
};
const xmlOf = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');

describe('line numbering', () => {
	it('reads w:lnNumType values', async () => {
		const loaded = await load(
			'<w:sectPr><w:lnNumType w:countBy="5" w:start="3" w:distance="360" w:restart="newPage"/></w:sectPr>',
		);
		expect(loaded.model.sections![0]).toMatchObject({
			lineNumbering: true,
			lineNumberSettings: { countBy: 5, start: 3, restart: 'newPage', distanceTwips: 360 },
		});
	});

	it('writes, changes and removes line numbers', async () => {
		const loaded = await load('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>');
		const edit = (change: (m: DocumentModel) => void) => {
			const model = structuredClone(loaded.model);
			change(model);
			return model;
		};
		const on = await xmlOf(
			await loaded.save(
				edit((m) => {
					m.sections![0]!.lineNumbering = true;
					m.sections![0]!.lineNumberSettings = { countBy: 1, start: 1, restart: 'newPage' };
				}),
			),
		);
		expect(on).toMatch(/<w:lnNumType [^>]*w:restart="newPage"/);
		expect(on.indexOf('lnNumType')).toBeGreaterThan(on.indexOf('pgSz'));

		const withLine = await load('<w:sectPr><w:lnNumType w:countBy="2"/></w:sectPr>');
		const removed = structuredClone(withLine.model);
		delete removed.sections![0]!.lineNumbering;
		delete removed.sections![0]!.lineNumberSettings;
		expect(await xmlOf(await withLine.save(removed))).not.toContain('lnNumType');
	});

	it('rejects a zero count before saving', async () => {
		const loaded = await load('<w:sectPr/>');
		const model = structuredClone(loaded.model);
		model.sections![0]!.lineNumberSettings = { countBy: 0, start: 1, restart: 'continuous' };
		await expect(loaded.save(model)).rejects.toThrow(/countBy/);
	});
});
