import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { ensureListDefinition } from './numbering-editing.js';
import type { Paragraph } from './model.js';
import { at, must } from './test-support/access.js';

const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const documentXml = `<?xml version="1.0"?><w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`;

async function fixtureWithoutNumbering(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file('word/document.xml', documentXml);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

async function fixtureWithNumbering(): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<?xml version="1.0"?><w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item one</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	zip.file(
		'word/numbering.xml',
		`<?xml version="1.0"?><w:numbering xmlns:w="${WORD_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
	);
	zip.file(
		'[Content_Types].xml',
		'<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
	);
	zip.file(
		'word/_rels/document.xml.rels',
		'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

describe('numbering package surgery', () => {
	it('parses numPr and numbering.xml into the model, and round-trips a no-op save unchanged', async () => {
		const bytes = await fixtureWithNumbering();
		const loaded = await loadDocx(bytes);
		const paragraph = loaded.model.blocks[0] as Paragraph;
		expect(paragraph.numbering).toEqual({ numId: 1, level: 0 });
		expect(loaded.model.numberingCatalog?.nums['1']).toMatchObject({ abstractNumId: '0' });
		expect(await loaded.save()).toEqual(bytes);
	});

	it('creates word/numbering.xml, its content-type and relationship the first time a document gains numbering', async () => {
		const bytes = await fixtureWithoutNumbering();
		const loaded = await loadDocx(bytes);
		const { catalog, numId } = ensureListDefinition(loaded.model.numberingCatalog, 'bullet');
		const paragraph = loaded.model.blocks[0] as Paragraph;
		paragraph.numbering = { numId, level: 0 };
		const saved = { ...loaded.model, numberingCatalog: catalog };
		const output = await JSZip.loadAsync(await loaded.save(saved));
		const numberingXml = await output.file('word/numbering.xml')?.async('string');
		expect(numberingXml).toContain(`w:numId="${numId}"`);
		expect(numberingXml).toContain('w:numFmt w:val="bullet"');
		const contentTypes = await output.file('[Content_Types].xml')?.async('string');
		expect(contentTypes).toContain('/word/numbering.xml');
		const rels = await output.file('word/_rels/document.xml.rels')?.async('string');
		expect(rels).toContain('numbering.xml');
		expect(rels).toContain(
			'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
		);
		// Reopening must resolve the newly created list.
		const reopened = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		expect((reopened.model.blocks[0] as Paragraph).numbering).toEqual({ numId, level: 0 });
	});

	it('appends a new list alongside an existing one without disturbing the original entries', async () => {
		const bytes = await fixtureWithNumbering();
		const loaded = await loadDocx(bytes);
		const { catalog, numId } = ensureListDefinition(loaded.model.numberingCatalog, 'decimal');
		expect(numId).not.toBe(1);
		const extra: Paragraph = {
			type: 'paragraph',
			id: 'p-extra',
			runs: [{ text: 'Second list' }],
			numbering: { numId, level: 0 },
		};
		const saved = {
			...loaded.model,
			blocks: [...loaded.model.blocks, extra],
			numberingCatalog: catalog,
		};
		const output = await JSZip.loadAsync(await loaded.save(saved));
		const numberingXml = (await output.file('word/numbering.xml')?.async('string')) ?? '';
		expect(numberingXml).toContain('w:abstractNumId="0"');
		expect(numberingXml).toContain(`w:numId="${numId}"`);
		expect(numberingXml.indexOf('<w:num ')).toBeGreaterThan(
			numberingXml.lastIndexOf('<w:abstractNum '),
		);
	});

	it('rejects edits to an existing numbering definition instead of silently rewriting it', async () => {
		const bytes = await fixtureWithNumbering();
		const loaded = await loadDocx(bytes);
		const catalog = loaded.model.numberingCatalog!;
		const abstract = must(catalog.abstractNums['0'], 'abstract numbering 0');
		const mutated = {
			...catalog,
			abstractNums: {
				...catalog.abstractNums,
				'0': {
					...abstract,
					levels: {
						...abstract.levels,
						0: { ...at(abstract.levels, 0), numFmt: 'upperRoman' as const },
					},
				},
			},
		};
		await expect(loaded.save({ ...loaded.model, numberingCatalog: mutated })).rejects.toThrow(
			'Cannot edit numbering definition',
		);
	});

	it('rejects creating a numbering catalog from the standalone writer', async () => {
		const { saveDocx } = await import('./save.js');
		const { catalog } = ensureListDefinition(undefined, 'bullet');
		await expect(
			saveDocx({
				blocks: [{ type: 'paragraph', id: 'p', runs: [{ text: 'x' }] }],
				page: {
					width: 816,
					height: 1056,
					marginTop: 96,
					marginRight: 96,
					marginBottom: 96,
					marginLeft: 96,
				},
				warnings: [],
				numberingCatalog: catalog,
			}),
		).rejects.toThrow('standalone DOCX writer');
	});
});
