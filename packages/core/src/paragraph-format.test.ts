import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { createDocument, loadDocx, saveDocx } from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const sourceXml = `<w:document xmlns:w="${ns}"><w:body>
<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:line="276" w:lineRule="auto" w:beforeAutospacing="1" w:afterAutospacing="1" w:customSpacing="keep"/><w:keepNext/></w:pPr><w:r><w:rPr><w:b w:val="off"/><w:i w:val="on"/><w:u w:val="0"/></w:rPr><w:t>Auto and left</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:line="360" w:lineRule="exact" w:beforeAutospacing="1"/><w:ind w:left="0"/></w:pPr><w:r><w:rPr><w:strike w:val="false"/></w:rPr><w:t>Exact</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:line="240" w:lineRule="atLeast" w:afterAutospacing="1" w:customSpacing="preserve"/></w:pPr><w:r><w:t>At least</w:t></w:r></w:p>
<w:p><w:pPr><w:spacing w:before="120" w:beforeAutospacing="1"/></w:pPr><w:r><w:t>Inherit</w:t></w:r></w:p>
<w:sectPr/></w:body></w:document>`;

async function loadedFixture() {
	const zip = new JSZip();
	zip.file('word/document.xml', sourceXml);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return { bytes, loaded: await loadDocx(bytes) };
}

async function documentXml(bytes: Uint8Array): Promise<string> {
	const zip = await JSZip.loadAsync(bytes);
	return (await zip.file('word/document.xml')?.async('string')) ?? '';
}

describe('DOCX paragraph formatting', () => {
	it('roundtrips a rule-only override without inventing an explicit line amount', async () => {
		const model = createDocument();
		const paragraph = model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.lineSpacingRule = 'atLeast';
		const loaded = await loadDocx(await saveDocx(model));
		const reopened = loaded.model.blocks[0];
		if (reopened.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(reopened.lineSpacingRule).toBe('atLeast');
		expect(reopened.lineSpacingTwips).toBeUndefined();
		reopened.spacingBeforeTwips = 100;
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('w:lineRule="atLeast"');
		expect(xml).not.toContain('w:line=');
	});
	it('reads explicit left alignment, line spacing modes, and on/off run properties', async () => {
		const { loaded } = await loadedFixture();
		const [auto, exact, atLeast, inherit] = loaded.model.blocks;
		if (
			auto?.type !== 'paragraph' ||
			exact?.type !== 'paragraph' ||
			atLeast?.type !== 'paragraph' ||
			inherit?.type !== 'paragraph'
		)
			throw new Error('Expected paragraph fixtures');
		expect(auto).toMatchObject({ align: 'left', lineSpacingTwips: 276, lineSpacingRule: 'auto' });
		expect(auto.runs[0]).toMatchObject({ italic: true });
		expect(auto.runs[0].bold).toBeUndefined();
		expect(auto.runs[0].underline).toBeUndefined();
		expect(exact).toMatchObject({
			lineSpacingTwips: 360,
			lineSpacingRule: 'exact',
			indentLeftTwips: 0,
		});
		expect(exact.runs[0].strike).toBeUndefined();
		expect(atLeast).toMatchObject({ lineSpacingTwips: 240, lineSpacingRule: 'atLeast' });
		expect(inherit.lineSpacingTwips).toBeUndefined();
		expect(inherit.lineSpacingRule).toBeUndefined();
	});

	it('returns byte-exact source for a no-op, including inherited and unsupported spacing values', async () => {
		const { bytes, loaded } = await loadedFixture();
		expect(await loaded.save()).toEqual(bytes);
	});

	it('updates line spacing modes while preserving unrelated spacing attributes and children', async () => {
		const { loaded } = await loadedFixture();
		const [auto, exact, atLeast, inherit] = loaded.model.blocks;
		if (
			auto?.type !== 'paragraph' ||
			exact?.type !== 'paragraph' ||
			atLeast?.type !== 'paragraph' ||
			inherit?.type !== 'paragraph'
		)
			throw new Error('Expected paragraph fixtures');
		auto.lineSpacingTwips = 480;
		auto.lineSpacingRule = 'exact';
		exact.lineSpacingRule = 'atLeast';
		atLeast.lineSpacingTwips = undefined;
		atLeast.lineSpacingRule = undefined;
		inherit.lineSpacingTwips = 300;
		inherit.lineSpacingRule = 'auto';
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('w:line="480" w:lineRule="exact"');
		expect(xml).toContain('w:line="360" w:lineRule="atLeast"');
		expect(xml).toContain('w:customSpacing="keep"');
		expect(xml).toContain('w:beforeAutospacing="1"');
		expect(xml).toContain('w:afterAutospacing="1"');
		expect(xml).not.toContain('w:line="240"');
		expect(xml).toContain('w:line="300" w:lineRule="auto"');
		expect(xml).toContain('<w:keepNext');
		const reopened = await loadDocx(await loaded.save());
		expect(reopened.model.blocks[0]).toMatchObject({
			lineSpacingTwips: 480,
			lineSpacingRule: 'exact',
		});
		expect(reopened.model.blocks[1]).toMatchObject({
			lineSpacingTwips: 360,
			lineSpacingRule: 'atLeast',
		});
		expect(reopened.model.blocks[2]).toMatchObject({});
		if (reopened.model.blocks[2]?.type !== 'paragraph') throw new Error('Expected paragraph');
		expect(reopened.model.blocks[2].spacingAfterTwips).toBeUndefined();
		expect(reopened.model.blocks[2].lineSpacingTwips).toBeUndefined();
		expect(reopened.model.blocks[3]).toMatchObject({
			lineSpacingTwips: 300,
			lineSpacingRule: 'auto',
		});
	});

	it('writes explicit left alignment and removes it when cleared', async () => {
		const { loaded } = await loadedFixture();
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.spacingBeforeTwips = 120;
		let xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:jc w:val="left"');
		paragraph.align = undefined;
		xml = await documentXml(await loaded.save());
		expect(xml).not.toContain('<w:jc w:val="left"');
	});

	it.each([
		'<w:br w:type="page"/>',
		'<w:br w:type="column"/>',
		'<w:br w:clear="left"/>',
		'<w:br w:unknown="keep"/>',
	])(
		'rejects paragraph edits that would flatten special break %s and preserves no-op bytes',
		async (breakXml) => {
			const zip = new JSZip();
			zip.file(
				'word/document.xml',
				`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:t>Before</w:t>${breakXml}<w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
			);
			const bytes = await zip.generateAsync({ type: 'uint8array' });
			const loaded = await loadDocx(bytes);
			expect(await loaded.save()).toEqual(bytes);
			expect(loaded.model.warnings).toContain(
				'Page, column, and other non-line breaks are not distinguished from line breaks in the document model; edits to paragraphs containing them are rejected to preserve the original XML.',
			);
			const paragraph = loaded.model.blocks[0];
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			paragraph.spacingBeforeTwips = 120;
			await expect(loaded.save()).rejects.toThrow('contains inline OOXML');
		},
	);

	it('allows paragraph formatting edits around an ordinary text-wrapping break', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:t>Before</w:t><w:br w:type="textWrapping"/><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
		paragraph.spacingBeforeTwips = 120;
		const xml = await documentXml(await loaded.save());
		expect(xml).toContain('<w:br/>');
		expect(xml).toContain('w:before="120"');
	});
});
