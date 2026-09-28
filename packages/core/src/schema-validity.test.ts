import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	buildTableOfContents,
	createDocument,
	ensureListDefinition,
	loadDocx,
	saveDocx,
	tocBookmarks,
	tocEntries,
	withTocBookmarks,
	type DocumentModel,
	type Paragraph,
} from './index.js';
import {
	packageSchemaErrors,
	schemaErrors,
	withoutExtensions,
} from './test-support/schema-validation.js';

function kitchenSink(): DocumentModel {
	const model = createDocument();
	const rev = { author: 'Ada', date: '2024-01-01T00:00:00Z' };
	model.blocks = [
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Heading' }], keepNext: true },
		{
			type: 'paragraph',
			id: 'p1',
			align: 'justify',
			spacingBeforeTwips: 120,
			lineSpacingTwips: 360,
			lineSpacingRule: 'auto',
			indentLeftTwips: 720,
			hangingTwips: 360,
			tabStops: [{ posTwips: 4680, align: 'right', leader: 'dot' }],
			pageBreakBefore: true,
			direction: 'rtl',
			runs: [
				{
					text: 'Bold',
					bold: true,
					italic: false,
					underline: true,
					strike: false,
					color: '#FF0000',
					fontSize: 14,
					fontFamily: 'Arial',
					highlight: 'yellow',
					verticalAlign: 'superscript',
					language: 'en-US',
					rtl: true,
					caps: true,
					smallCaps: false,
					vanish: false,
					characterSpacingTwips: 20,
					shadingFill: '#FFFF00',
					underlineStyle: 'double',
					underlineColor: '#00FF00',
					doubleStrike: false,
					style: 'Hyperlink',
				},
				{ text: 'link', link: { anchor: '_Toc100000000' } },
				{ text: '', break: 'page' },
				{ text: '\tafter tab\n' },
				{ text: 'inserted', revision: { ...rev, kind: 'insert', id: '5' } },
				{ text: 'deleted', revision: { ...rev, kind: 'delete', id: '6' } },
				{ text: 'moved', revision: { ...rev, kind: 'moveFrom', id: '7', move: { name: 'move1' } } },
				{ text: 'moved', revision: { ...rev, kind: 'moveTo', id: '8', move: { name: 'move1' } } },
				{ text: 'commented', commentIds: ['c1'] },
				{ text: '', noteReference: { kind: 'footnote', id: '1' } },
				{ text: '', fieldChar: 'begin' },
				{ text: '', fieldCode: ' PAGE ' },
				{ text: '', fieldChar: 'separate' },
				{ text: '1', field: { instr: 'PAGE' } },
				{ text: '', fieldChar: 'end' },
				{ text: 'Ann', field: { instr: 'AUTHOR', simple: true } },
			],
			markRevision: { ...rev, kind: 'insert', id: '9' },
		},
		{
			type: 'table',
			id: 't1',
			rows: [
				[
					{
						paragraphs: [{ type: 'paragraph', id: 'c1p', runs: [{ text: 'A' }] }],
						gridSpan: 1,
						verticalAlign: 'center',
						shadingFill: 'D9D9D9',
						widthTwips: 2000,
					},
					{ paragraphs: [{ type: 'paragraph', id: 'c2p', runs: [{ text: 'B' }] }] },
				],
			],
		},
		{ type: 'paragraph', id: 'end', runs: [{ text: 'End' }] },
	];
	model.comments = [{ id: 'c1', author: 'Ada', text: 'Note this', resolved: true }];
	model.footnotes = [
		{
			id: '1',
			blocks: [
				{
					type: 'paragraph',
					id: 'n1',
					style: 'FootnoteText',
					runs: [
						{ text: '', noteMark: 'footnote', style: 'FootnoteReference' },
						{ text: ' Source' },
					],
				},
			],
		},
	];
	return model;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
const PNG = Uint8Array.from(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
		'base64',
	),
);
const SVG = new TextEncoder().encode(
	'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>',
);

const clone = (model: DocumentModel) => structuredClone(model);
const paragraphs = (model: DocumentModel) =>
	model.blocks.filter((block): block is Paragraph => block.type === 'paragraph');

function sinkWithToc(): DocumentModel {
	const model = kitchenSink();
	const { bookmarks, added } = tocBookmarks(model, tocEntries(model));
	const marked = withTocBookmarks(model, added);
	let next = 0;
	marked.blocks.unshift(
		...buildTableOfContents(marked, {
			newId: () => `toc${++next}`,
			bookmarks,
			pageNumbers: new Map([['h1', '1']]),
		}),
	);
	return marked;
}

async function packageWith(files: Record<string, string>): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const [name, contents] of Object.entries(files)) zip.file(name, contents);
	return zip.generateAsync({ type: 'uint8array' });
}
const rels = (entries: string) =>
	`<?xml version="1.0"?><Relationships xmlns="${REL_PKG}">${entries}</Relationships>`;
const document = (body: string) =>
	`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`;

interface Fixture {
	name: string;
	build: () => Promise<Uint8Array>;
	/** Parts that must have been validated, so a fixture cannot silently skip one. */
	expected: string[];
}

const CORE = ['word/document.xml', 'word/styles.xml'];

const fixtures: Fixture[] = [
	{
		name: 'a new document using every supported feature',
		build: async () => saveDocx(sinkWithToc()),
		expected: [...CORE, 'word/comments.xml', 'word/footnotes.xml'],
	},
	{
		name: 'a loaded document that is edited and given a new list',
		build: async () => {
			const loaded = await loadDocx(await saveDocx(sinkWithToc()));
			const model = clone(loaded.model);
			const { catalog, numId } = ensureListDefinition(model.numberingCatalog, 'decimal');
			model.numberingCatalog = catalog;
			const plain = paragraphs(model).filter((paragraph) => !paragraph.style);
			plain[0].runs[0].text = 'Edited text';
			plain.at(-1)!.numbering = { numId, level: 0 };
			return loaded.save(model);
		},
		expected: [...CORE, 'word/comments.xml', 'word/footnotes.xml', 'word/numbering.xml'],
	},
	{
		name: 'headers and footers shared by two sections, edited',
		build: async () => {
			const refs = `<w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>`;
			const bytes = await packageWith({
				'word/document.xml': document(
					`<w:p><w:pPr><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr>${refs}<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`,
				),
				'word/_rels/document.xml.rels': rels(
					`<Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/><Relationship Id="rId2" Type="${R}/footer" Target="footer1.xml"/>`,
				),
				'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Draft</w:t></w:r></w:p></w:hdr>`,
				'word/footer1.xml': `<w:ftr xmlns:w="${W}"><w:p><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`,
			});
			const loaded = await loadDocx(bytes);
			const model = clone(loaded.model);
			for (const section of model.sections ?? [])
				(section.headers!.default!.blocks[0] as Paragraph).runs[0].text = 'Final';
			return loaded.save(model);
		},
		expected: ['word/document.xml', 'word/header1.xml', 'word/footer1.xml'],
	},
	{
		name: 'an endnote edited and a second endnote added',
		build: async () => {
			const bytes = await packageWith({
				'word/document.xml': document(
					`<w:p><w:r><w:t>Claim</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p><w:sectPr/>`,
				),
				'word/endnotes.xml': `<w:endnotes xmlns:w="${W}"><w:endnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Source.</w:t></w:r></w:p></w:endnote></w:endnotes>`,
			});
			const loaded = await loadDocx(bytes);
			const model = clone(loaded.model);
			(model.endnotes![0].blocks[0] as Paragraph).runs[1].text = ' Edited source.';
			model.endnotes!.push({
				id: '3',
				blocks: [
					{
						type: 'paragraph',
						id: 'en3',
						runs: [{ text: '', noteMark: 'endnote' }, { text: ' Second.' }],
					},
				],
			});
			paragraphs(model)[0].runs.push({ text: '', noteReference: { kind: 'endnote', id: '3' } });
			return loaded.save(model);
		},
		expected: ['word/document.xml', 'word/endnotes.xml'],
	},
	{
		name: 'a new picture beside a resized SVG picture',
		build: async () => {
			const first = createDocument();
			first.blocks[0] = {
				type: 'paragraph',
				id: 'p1',
				runs: [
					{
						text: '',
						image: {
							relId: '',
							partName: 'word/media/p.png',
							svgPartName: 'word/media/p.svg',
							contentType: 'image/png',
							widthPx: 40,
							heightPx: 40,
						},
					},
				],
			};
			const bytes = await saveDocx(
				first,
				new Map([
					['word/media/p.png', { bytes: PNG, contentType: 'image/png' }],
					['word/media/p.svg', { bytes: SVG, contentType: 'image/svg+xml' }],
				]),
			);
			const loaded = await loadDocx(bytes);
			const model = clone(loaded.model);
			const [paragraph] = paragraphs(model);
			paragraph.runs[0].image!.widthPx = 80;
			paragraph.runs[0].image!.heightPx = 20;
			paragraph.runs.push({
				text: '',
				image: {
					relId: '',
					partName: 'word/media/n.png',
					contentType: 'image/png',
					widthPx: 16,
					heightPx: 16,
					altText: 'New',
				},
			});
			return loaded.save(
				model,
				new Map([['word/media/n.png', { bytes: PNG, contentType: 'image/png' }]]),
			);
		},
		expected: ['word/document.xml'],
	},
	{
		name: 'a new table with merged, shaded, bordered and aligned cells and a short row',
		build: async () => {
			const model = createDocument();
			const cellParagraph = (id: string) => ({
				type: 'paragraph' as const,
				id,
				runs: [{ text: id }],
			});
			model.blocks.push({
				type: 'table',
				id: 'new-table',
				widthTwips: 9000,
				alignment: 'right',
				justification: 'end',
				indentTwips: 100,
				cellMargins: { left: 108, right: 108 },
				borders: { top: { style: 'single', sizeEighthPoints: 4, themeColor: 'accent1' } },
				rowProperties: [{ heightTwips: 500, heightRule: 'exact', cantSplit: true, header: true }],
				rows: [
					[
						{
							paragraphs: [cellParagraph('m1')],
							gridSpan: 2,
							verticalMerge: 'restart',
							widthTwips: 6000,
							shadingFill: '#D9D9D9',
							shadingThemeFill: { token: 'accent2', tint: 0.5 },
							verticalAlign: 'center',
							borders: { bottom: { style: 'double', sizeEighthPoints: 6, color: '#112233' } },
							margins: { top: 10, left: 20, bottom: 30, right: 40 },
						},
						{ paragraphs: [cellParagraph('m2')], widthTwips: 3000 },
					],
					[
						{ paragraphs: [cellParagraph('m3')], verticalMerge: 'continue', gridSpan: 2 },
						{ paragraphs: [cellParagraph('m4')] },
					],
					[{ paragraphs: [cellParagraph('m5')] }, { paragraphs: [cellParagraph('m6')] }],
				],
			});
			return saveDocx(model);
		},
		expected: ['word/document.xml'],
	},
	{
		name: 'a section without pgSz or pgMar but with a header reference, plus a new section break',
		build: async () => {
			const bytes = await packageWith({
				'word/document.xml': document(
					`<w:p><w:r><w:t>One</w:t></w:r></w:p><w:p><w:r><w:t>Two</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:cols w:space="720"/><w:titlePg/></w:sectPr>`,
				),
				'word/_rels/document.xml.rels': rels(
					`<Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/>`,
				),
				'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
			});
			const loaded = await loadDocx(bytes);
			const model = clone(loaded.model);
			const [last] = model.sections!;
			model.sections = [
				{ ...structuredClone(last), endsAtBlockId: model.blocks[0].id },
				{ ...last, type: 'continuous', columns: { count: 2, spacingTwips: 720, equalWidth: true } },
			];
			return loaded.save(model);
		},
		expected: ['word/document.xml', 'word/header1.xml'],
	},
];

describe('package schema validity', () => {
	it.each(fixtures)(
		'$name validates against ECMA-376',
		async ({ build, expected }) => {
			const { validated, errors } = await packageSchemaErrors(await build());
			expect(validated).toEqual(expect.arrayContaining(expected));
			expect(errors).toEqual({});
		},
		60000,
	);

	it('reports errors for a part that breaks the schema', async () => {
		const bogus = await packageWith({ 'word/document.xml': document('<w:p><w:bogus/></w:p>') });
		const { validated, errors } = await packageSchemaErrors(bogus);
		expect(validated).toEqual(['word/document.xml']);
		expect(Object.keys(errors)).toEqual(['word/document.xml']);
		const misordered = document('<w:p><w:pPr><w:jc w:val="left"/><w:keepNext/></w:pPr></w:p>');
		expect(await schemaErrors(misordered)).not.toEqual([]);
	});

	it('strips only extensions declared in mc:Ignorable', async () => {
		const w14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
		const paragraph = '<w:body><w:p w14:paraId="1"/></w:body>';
		const wrap = (attrs: string, inner: string) =>
			`<w:document xmlns:w="${W}" ${w14} ${MC_NS} ${attrs}>${inner}</w:document>`;
		expect(await schemaErrors(wrap('mc:Ignorable="w14"', paragraph))).toEqual([]);
		expect(await schemaErrors(wrap('', paragraph))).not.toEqual([]);
		const unknown = `<w:document xmlns:w="${W}" xmlns:x="urn:x"><w:body><x:thing/></w:body></w:document>`;
		expect(await schemaErrors(unknown)).not.toEqual([]);
	});

	it('unwraps AlternateContent to its Fallback and records a missing Fallback', () => {
		const choice = '<mc:Choice Requires="x"><w:x/></mc:Choice>';
		const wrap = (inner: string) =>
			`<w:body xmlns:w="${W}" ${MC_NS}><mc:AlternateContent>${inner}</mc:AlternateContent></w:body>`;
		const withFallback = withoutExtensions(wrap(`${choice}<mc:Fallback><w:p/></mc:Fallback>`));
		expect(withFallback.errors).toEqual([]);
		expect(withFallback.xml).toContain('<w:p/>');
		expect(withFallback.xml).not.toContain('w:x');
		expect(withoutExtensions(wrap(choice)).errors).toEqual([
			'mc:AlternateContent has no mc:Fallback',
		]);
	});
});
