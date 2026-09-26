import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const unicodeText = 'مرحبا עברית 中文 e\u0301 😀';
const sourceXml = `<w:document xmlns:w="${ns}"><w:body>
<w:p><w:pPr><w:bidi/><w:keepNext/></w:pPr><w:r><w:rPr><w:lang w:val="ar-SA" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:rtl/><w:i/></w:rPr><w:t>${unicodeText}</w:t></w:r></w:p>
<w:p><w:pPr><w:bidi w:val="off"/></w:pPr><w:r><w:rPr><w:lang w:val="he-IL"/><w:rtl w:val="0"/><w:highlight w:val="yellow"/></w:rPr><w:t>Second</w:t></w:r></w:p>
<w:p><w:r><w:t>Inherited direction</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

async function fixture() {
	const zip = new JSZip();
	zip.file('word/document.xml', sourceXml);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return { bytes, loaded: await loadDocx(bytes) };
}

async function xmlOf(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

describe('DOCX multilingual metadata', () => {
	it('reads direct paragraph direction, script languages, run RTL and Unicode verbatim', async () => {
		const { loaded } = await fixture();
		const [rtl, ltr, inherited] = loaded.model.blocks;
		if (rtl?.type !== 'paragraph' || ltr?.type !== 'paragraph' || inherited?.type !== 'paragraph')
			throw new Error('Expected paragraph fixtures');
		expect(rtl.direction).toBe('rtl');
		expect(rtl.runs[0]).toMatchObject({
			text: unicodeText,
			language: 'ar-SA',
			eastAsiaLanguage: 'ja-JP',
			bidiLanguage: 'ar-SA',
			rtl: true,
			italic: true,
		});
		expect(ltr.direction).toBe('ltr');
		expect(ltr.runs[0]).toMatchObject({ language: 'he-IL', rtl: false, highlight: 'yellow' });
		expect(inherited.direction).toBeUndefined();
		expect(inherited.runs[0].rtl).toBeUndefined();
	});

	it('returns original bytes untouched and edits only selected language and direction attributes', async () => {
		const { bytes, loaded } = await fixture();
		expect(await loaded.save()).toEqual(bytes);
		const [first, second] = loaded.model.blocks;
		if (first?.type !== 'paragraph' || second?.type !== 'paragraph')
			throw new Error('Expected paragraphs');
		first.direction = 'ltr';
		first.runs[0].language = 'en-NZ';
		first.runs[0].eastAsiaLanguage = undefined;
		first.runs[0].rtl = false;
		second.direction = 'rtl';
		const xml = await xmlOf(await loaded.save());
		expect(xml).toContain('<w:bidi w:val="0"');
		expect(xml).toContain('<w:bidi w:val="1"');
		expect(xml).toContain('<w:lang w:val="en-NZ" w:bidi="ar-SA"');
		expect(xml).not.toContain('w:eastAsia="ja-JP"');
		expect(xml).toContain('<w:rtl w:val="0"');
		expect(xml).toContain('<w:i');
		expect(xml).toContain('<w:highlight w:val="yellow"');
		const reopened = await loadDocx(await loaded.save());
		expect(reopened.model.blocks[0]).toMatchObject({
			direction: 'ltr',
			runs: [{ language: 'en-NZ', bidiLanguage: 'ar-SA', rtl: false, italic: true }],
		});
		expect(reopened.model.blocks[1]).toMatchObject({ direction: 'rtl', runs: [{ rtl: false }] });
	});

	it('validates newly written language tags and accepts script subtags', async () => {
		const { loaded } = await fixture();
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
		paragraph.runs[0].eastAsiaLanguage = 'zh-Hans-CN';
		expect(await loaded.save()).toBeInstanceOf(Uint8Array);
		paragraph.runs[0].language = 'en--US';
		await expect(loaded.save()).rejects.toThrow('Invalid BCP 47 language tag: en--US');
		paragraph.runs[0].language = 'en"><evil';
		await expect(loaded.save()).rejects.toThrow('Invalid BCP 47 language tag');
	});

	it('preserves supported language and explicit RTL-off metadata when a formatted run is split', async () => {
		const { loaded } = await fixture();
		const paragraph = loaded.model.blocks[1];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		const run = paragraph.runs[0];
		paragraph.runs = [
			{ ...run, text: 'Sec' },
			{ ...run, text: 'ond', bold: true },
		];
		const reloaded = await loadDocx(await loaded.save());
		expect(reloaded.model.blocks[1]).toMatchObject({
			runs: [
				{ text: 'Sec', language: 'he-IL', rtl: false },
				{ text: 'ond', language: 'he-IL', rtl: false, bold: true },
			],
		});
	});
});
