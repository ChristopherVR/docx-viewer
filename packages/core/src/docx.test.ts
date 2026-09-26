import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, saveDocx } from './index.js';

const sourceXml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:keepNext/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/document.xml', sourceXml);
	zip.file('word/styles.xml', '<styles><style id="custom"/></styles>');
	zip.file('customXml/item1.xml', '<payload>retain me</payload>');
	return zip.generateAsync({ type: 'uint8array' });
}

describe('DOCX core', () => {
	it('parses text, character formatting, alignment and page geometry', async () => {
		const loaded = await loadDocx(await fixture());
		const paragraph = loaded.model.blocks[0];
		expect(paragraph).toMatchObject({
			type: 'paragraph',
			align: 'center',
			runs: [{ text: 'Hello', bold: true }, { text: ' world' }],
		});
		expect(loaded.model.page).toMatchObject({ width: 816, height: 1056, marginTop: 96 });
	});

	it('reads and writes paragraph spacing and indents in native Word twips', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="exact" w:beforeAutospacing="1"/><w:ind w:left="720" w:right="360" w:firstLine="240"/></w:pPr><w:r><w:t>Styled</w:t></w:r></w:p><w:p><w:pPr><w:keepLines/><w:spacing w:after="120"/><w:ind w:start="480"/></w:pPr><w:r><w:t>Untouched</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const styled = loaded.model.blocks[0];
		const untouched = loaded.model.blocks[1];
		if (styled.type !== 'paragraph' || untouched.type !== 'paragraph')
			throw new Error('Expected paragraph fixtures');
		expect(styled).toMatchObject({
			spacingBeforeTwips: 120,
			spacingAfterTwips: 240,
			lineSpacingTwips: 360,
			lineSpacingRule: 'exact',
			indentLeftTwips: 720,
			indentRightTwips: 360,
			firstLineTwips: 240,
		});
		styled.spacingAfterTwips = 480;
		styled.indentLeftTwips = 960;
		untouched.spacingAfterTwips = 240;
		const saved = await JSZip.loadAsync(await loaded.save());
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('w:after="480"');
		expect(xml).toContain('w:left="960"');
		expect(xml).toContain('w:before="120"');
		expect(xml).toContain('w:line="360"');
		expect(xml).toContain('w:lineRule="exact"');
		expect(xml).toContain('w:beforeAutospacing="1"');
		expect(xml).toContain('w:keepNext');
		expect(xml).toContain('w:start="480"');
		expect(xml).toContain('w:keepLines');
		expect(untouched.runs[0].text).toBe('Untouched');
		expect(untouched.indentStartTwips).toBe(480);
		const reopened = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		expect(reopened.model.blocks[0]).toMatchObject({
			type: 'paragraph',
			spacingBeforeTwips: 120,
			spacingAfterTwips: 480,
			lineSpacingTwips: 360,
			lineSpacingRule: 'exact',
			indentLeftTwips: 960,
		});
	});

	it('returns original package bytes when untouched and preserves all parts', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph fixture');
		paragraph.runs[0].text = 'Changed';
		const zip = await JSZip.loadAsync(await loaded.save());
		expect(await zip.file('word/styles.xml')?.async('string')).toContain('custom');
		expect(await zip.file('customXml/item1.xml')?.async('string')).toContain('retain me');
		const xml = (await zip.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('Changed');
		expect(xml).toContain('w:keepNext');
		expect(xml).toContain('w:b');
		expect(xml).not.toContain('undefined');
		const reopened = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const savedParagraph = reopened.model.blocks[0];
		if (savedParagraph.type !== 'paragraph') throw new Error('Expected a paragraph fixture');
		expect(savedParagraph.runs.map((run) => run.text).join('')).toBe('Changed world');
	});

	it('creates a valid package from the standalone save API', async () => {
		const model = {
			blocks: [{ type: 'paragraph' as const, id: 'p', runs: [{ text: 'New document' }] }],
			page: {
				width: 816,
				height: 1056,
				marginTop: 96,
				marginRight: 96,
				marginBottom: 96,
				marginLeft: 96,
			},
			warnings: [],
		};
		const reopened = await loadDocx(await saveDocx(model));
		expect(reopened.model.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'New document' }],
		});
	});

	it('keeps paragraph and table order and exposes table cell paragraphs', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		expect(loaded.model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'table',
			'paragraph',
		]);
		const table = loaded.model.blocks[1];
		if (table.type !== 'table') throw new Error('Expected a table fixture');
		expect(table.rows[0][0].paragraphs[0].runs[0].text).toBe('Cell');
		const last = loaded.model.blocks[2];
		if (last.type !== 'paragraph') throw new Error('Expected final paragraph');
		last.runs[0].text = 'Edited';
		const saved = await JSZip.loadAsync(await loaded.save());
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml.indexOf('<w:p')).toBeLessThan(xml.indexOf('<w:tbl'));
		expect(xml.indexOf('<w:tbl')).toBeLessThan(xml.indexOf('Edited'));
	});

	it('preserves untouched hyperlink XML and rejects unsafe edits inside it', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:hyperlink w:history="1"><w:r><w:t>Link</w:t></w:r></w:hyperlink></w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const plain = loaded.model.blocks[1];
		if (plain.type !== 'paragraph') throw new Error('Expected a paragraph fixture');
		plain.runs[0].text = 'Changed';
		const output = await JSZip.loadAsync(await loaded.save());
		expect(await output.file('word/document.xml')?.async('string')).toContain('<w:hyperlink');
		const linked = loaded.model.blocks[0];
		if (linked.type !== 'paragraph') throw new Error('Expected a paragraph fixture');
		linked.runs[0].text = 'Changed link';
		await expect(loaded.save()).rejects.toThrow('Cannot edit paragraph p0');
	});

	it('keeps run text, tabs, breaks and explicit bold removal in order', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t></w:r></w:p><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const first = loaded.model.blocks[0];
		const second = loaded.model.blocks[1];
		if (first.type !== 'paragraph' || second.type !== 'paragraph')
			throw new Error('Expected paragraph fixtures');
		expect(first.runs[0].text).toBe('A\tB\nC');
		first.runs[0].text = 'X\tY\nZ';
		second.runs[0].bold = false;
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).not.toMatch(/<w:b(?:\s|\/>|>)/);
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		const savedFirst = reopened.model.blocks[0];
		if (savedFirst.type !== 'paragraph') throw new Error('Expected a paragraph');
		expect(savedFirst.runs[0].text).toBe('X\tY\nZ');
	});

	it('roundtrips direct strike, highlight and vertical alignment and clears edited values', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:strike/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/><w:lang w:val="en-US"/></w:rPr><w:t>Formatted</w:t></w:r><w:r><w:t> Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph fixture');
		expect(paragraph.runs[0]).toMatchObject({
			text: 'Formatted',
			strike: true,
			highlight: 'yellow',
			verticalAlign: 'superscript',
		});
		paragraph.runs[0].strike = false;
		paragraph.runs[0].highlight = undefined;
		paragraph.runs[0].verticalAlign = 'subscript';
		paragraph.runs[1].strike = true;
		paragraph.runs[1].highlight = 'cyan';
		const saved = await JSZip.loadAsync(await loaded.save());
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml.match(/<w:strike(?:\s|\/>|>)/g)).toHaveLength(1);
		expect(xml).not.toContain('<w:highlight w:val="yellow"');
		expect(xml).toContain('<w:vertAlign w:val="subscript"');
		expect(xml).toContain('<w:highlight w:val="cyan"');
		expect(xml).toContain('<w:lang w:val="en-US"');
		const reopened = await loadDocx(await saved.generateAsync({ type: 'uint8array' }));
		const savedParagraph = reopened.model.blocks[0];
		if (savedParagraph.type !== 'paragraph') throw new Error('Expected a paragraph');
		expect(savedParagraph.runs[0]).toMatchObject({ text: 'Formatted', verticalAlign: 'subscript' });
		expect(savedParagraph.runs[0].strike).toBeUndefined();
		expect(savedParagraph.runs[0].highlight).toBeUndefined();
		expect(savedParagraph.runs[1]).toMatchObject({ strike: true, highlight: 'cyan' });
	});

	it('keeps unknown body nodes in place while blocks are reordered', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>One</w:t></w:r></w:p><w:customXml w:element="marker"><w:opaque/></w:customXml><w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr/></w:body></w:document>',
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		loaded.model.blocks.reverse();
		const output = await JSZip.loadAsync(await loaded.save());
		const xml = (await output.file('word/document.xml')?.async('string')) ?? '';
		expect(xml.indexOf('Two')).toBeLessThan(xml.indexOf('customXml'));
		expect(xml.indexOf('customXml')).toBeLessThan(xml.indexOf('<w:tbl'));
		expect(xml).toContain('<w:opaque');
	});

	it('builds each save from the original package and reads zero margins', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original</w:t></w:r></w:p><w:sectPr><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0"/></w:sectPr></w:body></w:document>',
		);
		const original = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(original);
		expect(loaded.model.page).toMatchObject({
			marginTop: 0,
			marginRight: 0,
			marginBottom: 0,
			marginLeft: 0,
		});
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
		paragraph.runs[0].text = 'First save';
		const firstSave = await loaded.save();
		paragraph.runs[0].text = 'Second save';
		const secondSave = await loaded.save();
		const firstZip = await JSZip.loadAsync(firstSave);
		const secondZip = await JSZip.loadAsync(secondSave);
		expect(await firstZip.file('word/document.xml')?.async('string')).toContain('First save');
		expect(await secondZip.file('word/document.xml')?.async('string')).toContain('Second save');
		const originalZip = await JSZip.loadAsync(original);
		expect(await originalZip.file('word/document.xml')?.async('string')).toContain('Original');
	});

	it('rejects malformed XML before exposing a partial document', async () => {
		const zip = new JSZip();
		zip.file('word/document.xml', '<w:document><w:body>');
		await expect(loadDocx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(
			'Invalid DOCX XML',
		);
	});
});
