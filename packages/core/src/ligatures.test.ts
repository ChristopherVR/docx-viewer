import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
	createDocument,
	loadDocx,
	saveDocx,
	resolveRunFormatting,
	type Ligatures,
} from './index.js';
import { LIGATURE_VALUES, WORD_2010_NS, writeLigatures } from './ligatures.js';
import { parseXml, WORD_NS } from './xml.js';
import { runHasUnknownProperties } from './write-run-validation.js';
import { parseRunProperties } from './run-properties.js';
import { expectParagraph } from './test-support/access.js';

const xmlRun = (value: string) =>
	parseXml(
		`<w:r xmlns:w="${WORD_NS}" xmlns:x="${WORD_2010_NS}"><w:rPr><x:ligatures ${value}/></w:rPr><w:t>office</w:t></w:r>`,
	).documentElement;
describe('Office 2010 ligatures', () => {
	it('preserves conflicting namespace bindings when adding extensions', () => {
		const doc = parseXml(
			`<w:rPr xmlns:w="${WORD_NS}" xmlns:w14="urn:existing" xmlns:mc="urn:other"/>`,
		);
		writeLigatures(doc, doc.documentElement, 'all');
		expect(doc.documentElement.lookupNamespaceURI('w14')).toBe('urn:existing');
		expect(doc.documentElement.lookupNamespaceURI('mc')).toBe('urn:other');
		expect(doc.documentElement.lookupNamespaceURI('w141')).toBe(WORD_2010_NS);
		expect(
			doc.documentElement.getAttributeNS(
				'http://schemas.openxmlformats.org/markup-compatibility/2006',
				'Ignorable',
			),
		).toBe('w141');
		expect(parseRunProperties(doc.documentElement).ligatures).toBe('all');
	});
	it.each(LIGATURE_VALUES)(
		'parses and saves %s with an ignorable extension declaration',
		async (value) => {
			const model = createDocument();
			model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'office', ligatures: value }] }];
			const bytes = await saveDocx(model);
			const zip = await JSZip.loadAsync(bytes);
			const xml = await zip.file('word/document.xml')!.async('string');
			expect(xml).toContain('mc:Ignorable="w14"');
			expect(xml).toContain(`w14:val="${value}"`);
			expect(expectParagraph((await loadDocx(bytes)).model.blocks[0]).runs[0]!.ligatures).toBe(
				value,
			);
			expect(runHasUnknownProperties(xmlRun(`x:val="${value}"`))).toBe(false);
		},
	);
	it('resolves inheritance and explicit None', () => {
		const context = {
			runCatalog: { docDefaults: { ligatures: 'all' as Ligatures }, styles: {}, warnings: [] },
		};
		expect(resolveRunFormatting({ text: 'office' }, context).ligatures).toBe('all');
		expect(resolveRunFormatting({ text: 'office', ligatures: 'none' }, context).ligatures).toBe(
			'none',
		);
	});
	it.each(['x:val="unknown"', '', 'x:val="all" x:custom="keep"', 'w:val="all"'])(
		'keeps invalid or unmodeled extension protected: %s',
		(attrs) => {
			expect(runHasUnknownProperties(xmlRun(attrs))).toBe(true);
		},
	);
	it('preserves imported surrounding extensions and removes/reset only ligatures', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${WORD_NS}" xmlns:w14="${WORD_2010_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w:r><w:rPr><w14:ligatures w14:val="standardContextual"/></w:rPr><w:t>office</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		);
		const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const run = expectParagraph(model.blocks[0]).runs[0]!;
		run.text = 'affinity';
		run.ligatures = 'none';
		expect(
			expectParagraph((await loadDocx(await saveDocx(model))).model.blocks[0]).runs[0],
		).toMatchObject({ text: 'affinity', ligatures: 'none' });
		delete run.ligatures;
		const output = await JSZip.loadAsync(await saveDocx(model));
		expect(await output.file('word/document.xml')!.async('string')).not.toContain('<w14:ligatures');
	});
	it('does not accept extension-looking elements from the wrong namespace', () => {
		const props = parseXml(
			`<w:rPr xmlns:w="${WORD_NS}"><w:ligatures w:val="all"/></w:rPr>`,
		).documentElement;
		expect(parseRunProperties(props).ligatures).toBeUndefined();
	});
});
