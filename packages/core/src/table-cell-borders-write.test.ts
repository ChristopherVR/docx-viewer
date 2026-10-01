import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Table } from './model.js';
import { eighthPoints } from './units.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (text: string, props = '') =>
	`<w:tc>${props}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell(
			'A',
			'<w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:left w:val="single" w:sz="8" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr>',
		)}${cell('B')}</w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

const save = async (mutate: (table: Table) => void) => {
	const loaded = await loadDocx(await fixture());
	mutate(loaded.model.blocks[0] as Table);
	const output = await JSZip.loadAsync(await loaded.save());
	return {
		xml: (await output.file('word/document.xml')?.async('string')) ?? '',
		reopened: await loadDocx(await output.generateAsync({ type: 'uint8array' })),
	};
};
const pen = { style: 'single' as const, sizeEighthPoints: eighthPoints(12), color: '#00ff00' };

describe('cell border edits', () => {
	it('changes only the edited side, keeping the other sides, width and shading, in schema order', async () => {
		const { xml, reopened } = await save((table) => {
			table.rows[0]![0]!.borders = { ...table.rows[0]![0]!.borders, bottom: pen };
		});
		expect(xml).toMatch(
			/<w:tcPr><w:tcW [^>]*\/><w:tcBorders><w:left [^>]*w:color="FF0000"[^>]*\/><w:bottom [^>]*w:sz="12"[^>]*w:color="00FF00"[^>]*\/><\/w:tcBorders><w:shd /,
		);
		const read = (reopened.model.blocks[0] as Table).rows[0]![0]!;
		expect(read.borders?.bottom?.sizeEighthPoints).toBe(12);
		expect(read.borders?.left?.color?.toLowerCase()).toBe('#ff0000');
		expect(read.shadingFill).toBeTruthy();
	});

	it('creates tcPr for a cell without one and removes borders it clears', async () => {
		const { xml } = await save((table) => {
			table.rows[0]![1]!.borders = { top: pen };
			delete table.rows[0]![0]!.borders;
		});
		expect(xml).toMatch(/<w:tc><w:tcPr><w:tcBorders><w:top [^>]*\/><\/w:tcBorders><\/w:tcPr><w:p>/);
		expect(xml).not.toContain('FF0000');
		expect(xml).toContain('w:fill="EEEEEE"');
	});

	it('leaves the package untouched when borders are unchanged', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
	});
});
