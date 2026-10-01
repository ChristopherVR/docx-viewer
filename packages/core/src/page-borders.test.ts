import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { eighthPoints } from './units.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

async function fixture(borders: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>${borders}<w:pgNumType w:start="1"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const xmlOf = async (bytes: Uint8Array) =>
	(await (await JSZip.loadAsync(bytes)).file('word/document.xml')?.async('string')) ?? '';
const WITH =
	'<w:pgBorders w:offsetFrom="text" w:display="firstPage"><w:top w:val="single" w:sz="8" w:space="24" w:color="FF0000"/><w:left w:val="single" w:sz="4" w:space="4" w:color="auto" w:art="apples"/></w:pgBorders>';

describe('page borders', () => {
	it('parses sides, spacing, art, offset and display', async () => {
		const { model } = await loadDocx(await fixture(WITH));
		expect(model.sections![0]!.pageBorders).toMatchObject({
			offsetFrom: 'text',
			display: 'firstPage',
			top: { style: 'single', sizeEighthPoints: 8, color: '#FF0000', spacePoints: 24 },
			left: { spacePoints: 4, art: 'apples' },
		});
		expect(model.sections![0]!.pageBorders!.bottom).toBeUndefined();
	});

	it('rewrites only the changed side and keeps the art side and attributes', async () => {
		const loaded = await loadDocx(await fixture(WITH));
		const borders = loaded.model.sections![0]!.pageBorders!;
		borders.top = { ...borders.top!, sizeEighthPoints: eighthPoints(18), color: '#00ff00' };
		borders.bottom = { style: 'double', sizeEighthPoints: eighthPoints(6), spacePoints: 10 };
		const xml = await xmlOf(await loaded.save());
		expect(xml).toMatch(
			/<w:pgBorders w:offsetFrom="text" w:display="firstPage"><w:top [^>]*w:sz="18"[^>]*w:color="00FF00"[^>]*\/><w:left [^>]*w:art="apples"[^>]*\/><w:bottom w:val="double" w:sz="6" w:space="10" w:color="auto"\/><\/w:pgBorders><w:pgNumType/,
		);
	});

	it('creates pgBorders in schema order and removes it again', async () => {
		const loaded = await loadDocx(await fixture(''));
		loaded.model.sections![0]!.pageBorders = {
			right: { style: 'single', sizeEighthPoints: eighthPoints(4), spacePoints: 24 },
			offsetFrom: 'page',
		};
		const saved = await loaded.save();
		const xml = await xmlOf(saved);
		expect(xml).toMatch(
			/<w:pgMar [^>]*\/><w:pgBorders w:offsetFrom="page"><w:right [^>]*\/><\/w:pgBorders><w:pgNumType/,
		);
		const reopened = await loadDocx(saved);
		delete reopened.model.sections![0]!.pageBorders;
		expect(await xmlOf(await reopened.save())).not.toContain('pgBorders');
	});

	it('keeps the package byte-exact when borders are untouched', async () => {
		const bytes = await fixture(WITH);
		expect(await (await loadDocx(bytes)).save()).toEqual(bytes);
	});
});
