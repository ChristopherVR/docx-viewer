import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { expectParagraph } from './test-support/access.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';

async function fixture(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Reviewed text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/comments.xml',
		`<w:comments ${NS} ${W14}><w:comment w:id="0" w:author="Ada" w:date="2024-01-01T00:00:00Z" w:initials="AL"><w:p w14:paraId="00000001"><w:r><w:t>Please check this.</w:t></w:r></w:p></w:comment></w:comments>`,
	);
	zip.file(
		'word/commentsExtended.xml',
		`<w15:commentsEx ${W15}><w15:commentEx w15:paraId="00000001" w15:done="0"/></w15:commentsEx>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('comments', () => {
	it('parses comment text/author and anchors the range onto the covered run', async () => {
		const loaded = await loadDocx(await fixture());
		expect(loaded.model.comments).toMatchObject([
			{ id: '0', author: 'Ada', initials: 'AL', text: 'Please check this.', resolved: false },
		]);
		const paragraph = expectParagraph(loaded.model.blocks[0]);
		expect(paragraph.runs[0]).toMatchObject({ text: 'Reviewed text', commentIds: ['0'] });
		expect(loaded.model.warnings.some((w) => w.includes('run granularity'))).toBe(true);
	});

	it('leaves an untouched comment list byte-identical on save', async () => {
		const bytes = await fixture();
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
	});

	it('adds a new comment, resolves an existing one, and creates comments.xml parts from scratch', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t>Plain text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		zip.file(
			'[Content_Types].xml',
			'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		expect(loaded.model.comments ?? []).toHaveLength(0);
		const next = {
			...loaded.model,
			comments: [{ id: 'c1', author: 'Ada', text: 'New comment', resolved: true }],
		};
		const saved = await JSZip.loadAsync(await loaded.save(next));
		const commentsXml = await saved.file('word/comments.xml')?.async('string');
		expect(commentsXml).toContain('New comment');
		const extendedXml = await saved.file('word/commentsExtended.xml')?.async('string');
		expect(extendedXml).toContain('w15:done="1"');
		const contentTypes = await saved.file('[Content_Types].xml')?.async('string');
		expect(contentTypes).toContain('comments+xml');
		const rels = await saved.file('word/_rels/document.xml.rels')?.async('string');
		expect(rels).toContain('/comments');
	});

	it('models a comment spanning paragraphs and writes one range for it', async () => {
		const zip = await JSZip.loadAsync(await fixture());
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:commentRangeStart w:id="0"/><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>middle</w:t></w:r></w:p><w:p><w:r><w:t>last</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const runs = loaded.model.blocks.map((block) => (block.type === 'paragraph' ? block.runs : []));
		expect(runs.map((list) => list.map((run) => run.commentIds ?? []))).toEqual([
			[[], ['0']],
			[['0']],
			[['0'], []],
		]);
		// Edit the middle and last paragraphs; the range still opens in the first and closes in the last.
		const blocks = loaded.model.blocks.map((block, index) =>
			block.type === 'paragraph' && index > 0
				? { ...block, runs: block.runs.map((run) => ({ ...run, text: run.text.toUpperCase() })) }
				: block,
		);
		const saved = await loaded.save({ ...loaded.model, blocks });
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml.match(/<w:commentRangeStart /g)).toHaveLength(1);
		expect(xml.match(/<w:commentRangeEnd /g)).toHaveLength(1);
		expect(xml.match(/<w:commentReference /g)).toHaveLength(1);
		expect(xml).toMatch(
			/<w:t>LAST<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/><w:r><w:commentReference w:id="0"\/>/,
		);
		// Moving the end: drop the comment from the last paragraph, and the middle one closes it.
		const shortened = blocks.map((block, index) =>
			block.type === 'paragraph' && index === 2
				? { ...block, runs: block.runs.map(({ commentIds: _ids, ...run }) => run) }
				: block,
		);
		const shortXml = await (
			await JSZip.loadAsync(await loaded.save({ ...loaded.model, blocks: shortened }))
		)
			.file('word/document.xml')!
			.async('string');
		expect(shortXml.match(/<w:commentRangeEnd /g)).toHaveLength(1);
		expect(shortXml).toMatch(/<w:t>MIDDLE<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/>/);
		const reloaded = await loadDocx(await loaded.save({ ...loaded.model, blocks: shortened }));
		const last = expectParagraph(reloaded.model.blocks[2]);
		expect(last.runs.every((run) => !run.commentIds)).toBe(true);
	});
});
