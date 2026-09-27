import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import {
	acceptRevision,
	listRevisions,
	rejectRevision,
	acceptAllRevisions,
} from './revision-commands.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function docx(bodyXml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('tracked-change revisions', () => {
	it('parses inserted and deleted runs with author/date/id', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:r><w:t>Hello </w:t></w:r>' +
					'<w:ins w:id="1" w:author="Ada" w:date="2024-01-01T00:00:00Z"><w:r><w:t>brave </w:t></w:r></w:ins>' +
					'<w:del w:id="2" w:author="Grace"><w:r><w:delText>old </w:delText></w:r></w:del>' +
					'<w:r><w:t>world</w:t></w:r></w:p>',
			),
		);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');
		expect(paragraph.runs).toMatchObject([
			{ text: 'Hello ' },
			{ text: 'brave ', revision: { kind: 'insert', author: 'Ada', id: '1' } },
			{ text: 'old ', revision: { kind: 'delete', author: 'Grace', id: '2' } },
			{ text: 'world' },
		]);
	});

	it('parses paragraph mark insertion and pPrChange/rPrChange markers, and reports them honestly', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:pPr><w:rPr><w:ins w:id="9" w:author="Ada"/></w:rPr></w:pPr><w:r><w:t>New para</w:t></w:r></w:p>' +
					'<w:p><w:pPr><w:pPrChange w:id="10" w:author="Ada"><w:pPr/></w:pPrChange></w:pPr>' +
					'<w:r><w:rPr><w:rPrChange w:id="11" w:author="Ada"><w:rPr/></w:rPrChange></w:rPr><w:t>Reformatted</w:t></w:r></w:p>',
			),
		);
		const [first, second] = loaded.model.blocks;
		if (first.type !== 'paragraph' || second.type !== 'paragraph')
			throw new Error('expected paragraphs');
		expect(first.markRevision).toMatchObject({ kind: 'insert', author: 'Ada', id: '9' });
		expect(second.formatRevision).toMatchObject({ kind: 'paragraphChange', id: '10' });
		expect(second.runs[0].revision).toMatchObject({ kind: 'formatChange', id: '11' });
		expect(loaded.model.warnings.some((w) => w.includes('Formatting-change revisions'))).toBe(true);
		expect(loaded.model.warnings.some((w) => w.includes('Comments and tracked review'))).toBe(
			false,
		);
	});

	it('parses moveFrom/moveTo as a delete/insert pair and warns about the missing move linkage', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:moveFrom w:id="3" w:author="A"><w:r><w:t>moved </w:t></w:r></w:moveFrom>' +
					'<w:moveTo w:id="4" w:author="A"><w:r><w:t>moved </w:t></w:r></w:moveTo><w:r><w:t>text</w:t></w:r></w:p>',
			),
		);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');
		expect(paragraph.runs[0].revision?.kind).toBe('moveFrom');
		expect(paragraph.runs[1].revision?.kind).toBe('moveTo');
		expect(loaded.model.warnings.some((w) => w.includes('move linkage'))).toBe(true);
	});

	it('round-trips an insertion revision through accept/reject/save', async () => {
		const bytes = await docx(
			'<w:p><w:r><w:t>Hello </w:t></w:r><w:ins w:id="1" w:author="Ada" w:date="2024-01-01T00:00:00Z"><w:r><w:t>brave </w:t></w:r></w:ins><w:r><w:t>world</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const revisions = listRevisions(loaded.model);
		expect(revisions).toHaveLength(1);
		const accepted = acceptRevision(loaded.model, revisions[0].id);
		const acceptedParagraph = accepted.blocks[0];
		if (acceptedParagraph.type !== 'paragraph') throw new Error('expected paragraph');
		expect(acceptedParagraph.runs.map((r) => r.text).join('')).toBe('Hello brave world');
		expect(acceptedParagraph.runs.every((r) => !r.revision)).toBe(true);
		const savedAccepted = await JSZip.loadAsync(await loaded.save(accepted));
		const acceptedXml = (await savedAccepted.file('word/document.xml')?.async('string')) ?? '';
		expect(acceptedXml).not.toContain('<w:ins');
		expect(acceptedXml).toContain('brave');

		const rejected = rejectRevision(loaded.model, revisions[0].id);
		const rejectedParagraph = rejected.blocks[0];
		if (rejectedParagraph.type !== 'paragraph') throw new Error('expected paragraph');
		expect(rejectedParagraph.runs.map((r) => r.text).join('')).toBe('Hello world');
		const savedRejected = await JSZip.loadAsync(await loaded.save(rejected));
		const rejectedXml = (await savedRejected.file('word/document.xml')?.async('string')) ?? '';
		expect(rejectedXml).not.toContain('brave');
	});

	it('round-trips a deletion revision, writing w:delText and restoring text on reject', async () => {
		const bytes = await docx(
			'<w:p><w:r><w:t>Hello </w:t></w:r><w:del w:id="2" w:author="Grace"><w:r><w:delText>old </w:delText></w:r></w:del><w:r><w:t>world</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const [revision] = listRevisions(loaded.model);
		const accepted = acceptRevision(loaded.model, revision.id);
		expect((accepted.blocks[0] as any).runs.map((r: any) => r.text).join('')).toBe('Hello world');
		const rejected = rejectRevision(loaded.model, revision.id);
		expect((rejected.blocks[0] as any).runs.map((r: any) => r.text).join('')).toBe(
			'Hello old world',
		);
		const savedRejected = await JSZip.loadAsync(await loaded.save(rejected));
		const xml = (await savedRejected.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).not.toContain('<w:del');
		expect(xml).toContain('old');

		const savedAccepted = await JSZip.loadAsync(await loaded.save(accepted));
		const acceptedXml = (await savedAccepted.file('word/document.xml')?.async('string')) ?? '';
		expect(acceptedXml).not.toContain('old');
	});

	it('re-serializes a pending deletion revision as w:del/w:delText when the paragraph is otherwise edited', async () => {
		const bytes = await docx(
			'<w:p><w:r><w:t>Hello </w:t></w:r><w:del w:id="2" w:author="Grace"><w:r><w:delText>old </w:delText></w:r></w:del><w:r><w:t>world</w:t></w:r></w:p>',
		);
		const loaded = await loadDocx(bytes);
		const paragraph = loaded.model.blocks[0];
		if (paragraph.type !== 'paragraph') throw new Error('expected paragraph');
		const next = structuredClone(loaded.model);
		(next.blocks[0] as typeof paragraph).runs[0].text = 'Hi ';
		const saved = await JSZip.loadAsync(await loaded.save(next));
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('<w:del ');
		expect(xml).toContain('<w:delText');
		expect(xml).toContain('old ');
		expect(xml).toContain('Hi ');
	});

	it('acceptAllRevisions clears every revision and keeps a no-op save returning original bytes', async () => {
		const bytes = await docx('<w:p><w:r><w:t>Plain</w:t></w:r></w:p>');
		const loaded = await loadDocx(bytes);
		expect(await loaded.save()).toEqual(bytes);
		const accepted = acceptAllRevisions(loaded.model);
		expect(accepted).toEqual(loaded.model);
	});

	it('rejects an edit that would merge a paragraph mark deletion across a table boundary', async () => {
		const loaded = await loadDocx(
			await docx(
				'<w:p><w:pPr><w:rPr><w:del w:id="5" w:author="A"/></w:rPr></w:pPr><w:r><w:t>Last para</w:t></w:r></w:p>',
			),
		);
		const [revision] = listRevisions(loaded.model);
		expect(() => acceptRevision(loaded.model, revision.id)).toThrow(/merging/);
	});
});
