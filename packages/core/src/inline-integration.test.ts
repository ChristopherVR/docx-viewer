import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type Paragraph } from './index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function fixture(body: string, rels: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
	);
	zip.file('word/styles.xml', `<w:styles xmlns:w="${W}"/>`);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const stylesRel = `<Relationship Id="rId1" Type="${REL_TYPE}/styles" Target="styles.xml"/>`;

async function partsOf(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	return {
		document: await zip.file('word/document.xml')!.async('string'),
		rels: await zip.file('word/_rels/document.xml.rels')!.async('string'),
	};
}

describe('inline content across features', () => {
	it('never reuses a relationship id that only the .rels part declares', async () => {
		const loaded = await loadDocx(
			await fixture('<w:p><w:r><w:t>Visit</w:t></w:r></w:p>', stylesRel),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.runs = [{ text: 'Visit', link: { href: 'https://example.com/' } }];
		const { document, rels } = await partsOf(await loaded.save(loaded.model));
		expect(rels).toContain('Target="styles.xml"');
		const linkId = /<w:hyperlink[^>]*r:id="(rId\d+)"/.exec(document)?.[1];
		expect(linkId).toBeDefined();
		expect(linkId).not.toBe('rId1');
		expect(rels).toContain(`Id="${linkId}"`);
		expect(rels).toContain('Target="https://example.com/"');
	});

	it('keeps a tracked insertion inside an existing hyperlink editable and preserved', async () => {
		const loaded = await loadDocx(
			await fixture(
				'<w:p><w:hyperlink r:id="rId2"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:ins w:id="7" w:author="Ann"><w:r><w:t> added</w:t></w:r></w:ins><w:r><w:t> tail</w:t></w:r></w:p>',
				`${stylesRel}<Relationship Id="rId2" Type="${REL_TYPE}/hyperlink" Target="https://a.example/" TargetMode="External"/>`,
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs[0]).toMatchObject({ text: 'Link', link: { href: 'https://a.example/' } });
		expect(paragraph.runs[1]).toMatchObject({
			text: ' added',
			revision: { kind: 'insert', author: 'Ann' },
		});
		paragraph.runs[2].text = ' changed tail';
		const { document } = await partsOf(await loaded.save(loaded.model));
		expect(document).toMatch(
			/<w:hyperlink r:id="rId2"[^>]*><w:r><w:t[^>]*>Link<\/w:t><\/w:r><\/w:hyperlink>/,
		);
		expect(document).toContain('w:author="Ann"');
		expect(document).toContain('changed tail');
	});

	it('does not duplicate comment reference runs across repeated saves', async () => {
		const loaded = await loadDocx(
			await fixture(
				'<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Noted</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t> text</w:t></w:r></w:p>',
				stylesRel,
			),
		);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.runs.map((run) => run.text)).toEqual(['Noted', ' text']);
		paragraph.runs[1].text = ' edited';
		const once = await loaded.save(loaded.model);
		const reloaded = await loadDocx(once);
		(reloaded.model.blocks[0] as Paragraph).runs[1].text = ' edited again';
		const { document } = await partsOf(await reloaded.save(reloaded.model));
		expect(document.match(/<w:commentReference/g)).toHaveLength(1);
		expect(document).toContain('edited again');
	});
});
