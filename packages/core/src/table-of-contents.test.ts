import { twips } from './units.js';
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import {
	buildTableOfContents,
	createDocument,
	findTableOfContents,
	headingLevel,
	loadDocx,
	tocLevels,
	updateTableOfContents,
	tocBookmarks,
	tocEntries,
	withTocBookmarks,
	saveDocx,
	type DocumentModel,
	type Paragraph,
} from './index.js';
import { at } from './test-support/access.js';

function sample(): DocumentModel {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'h1', style: 'Heading1', runs: [{ text: 'Introduction' }] },
		{ type: 'paragraph', id: 'b1', runs: [{ text: 'Body' }] },
		{ type: 'paragraph', id: 'h2', style: 'Heading2', runs: [{ text: 'Scope\tand aims' }] },
		{ type: 'paragraph', id: 'h4', style: 'Heading4', runs: [{ text: 'Too deep' }] },
	];
	return model;
}

let counter = 0;
const newId = () => `toc-${++counter}`;
const text = (paragraph: Paragraph) => paragraph.runs.map((run) => run.text).join('');

describe('table of contents', () => {
	it('detects heading levels from style ids, built-in names and basedOn chains', () => {
		const catalog = {
			docDefaults: {},
			warnings: [],
			styles: {
				Titre1: { id: 'Titre1', name: 'heading 1', formatting: {} },
				Custom: { id: 'Custom', name: 'My heading', basedOn: 'Titre1', formatting: {} },
			},
		};
		expect(headingLevel({ type: 'paragraph', id: 'a', runs: [], style: 'Heading3' })).toBe(3);
		expect(headingLevel({ type: 'paragraph', id: 'b', runs: [], style: 'Titre1' }, catalog)).toBe(
			1,
		);
		expect(headingLevel({ type: 'paragraph', id: 'c', runs: [], style: 'Custom' }, catalog)).toBe(
			1,
		);
		expect(headingLevel({ type: 'paragraph', id: 'd', runs: [] })).toBeUndefined();
		expect(tocLevels(' TOC \\o "2-4" \\h ')).toEqual({ from: 2, to: 4 });
	});

	it('builds entries with page numbers inside one field spanning the entry paragraphs', () => {
		const pageNumbers = new Map([
			['h1', '1'],
			['h2', '2'],
		]);
		const toc = buildTableOfContents(sample(), {
			newId,
			pageNumbers,
			contentWidthTwips: twips(9000),
		});
		expect(toc.map(text)).toEqual(['Introduction\t1', 'Scope and aims\t2']);
		expect(at(toc, 0).runs.slice(0, 3)).toEqual([
			{ text: '', fieldChar: 'begin' },
			{ text: '', fieldCode: ' TOC \\o "1-3" \\h \\z \\u ' },
			{ text: '', fieldChar: 'separate' },
		]);
		expect(at(toc, 1).runs.at(-1)).toEqual({ text: '', fieldChar: 'end' });
		// New documents define Word's TOC styles, so entries use them.
		expect(toc[1]).toMatchObject({
			style: 'TOC2',
			tabStops: [{ posTwips: 9000, align: 'right', leader: 'dot' }],
		});
		// Without TOC styles, entries get matching direct indents.
		const { paragraphStyles: _styles, ...unstyled } = sample();
		expect(buildTableOfContents(unstyled, { newId })[1]).toMatchObject({
			indentLeftTwips: 220,
			spacingAfterTwips: 100,
		});
	});

	it("shows Word's empty message when there are no headings", () => {
		const model = createDocument();
		const toc = buildTableOfContents(model, { newId });
		expect(toc).toHaveLength(1);
		expect(text(at(toc, 0))).toBe('No table of contents entries found.');
	});

	it('updates an existing TOC in place, keeping its switches and surrounding text', () => {
		const model = sample();
		const toc = buildTableOfContents(model, { newId, instruction: ' TOC \\o "1-1" ' });
		at(toc, 0).runs.unshift({ text: 'Lead ' });
		toc.at(-1)!.runs.push({ text: ' tail' });
		model.blocks.unshift(...toc);
		expect(findTableOfContents(model.blocks)).toMatchObject({ start: 0, end: 0 });
		model.blocks.push({
			type: 'paragraph',
			id: 'h1b',
			style: 'Heading1',
			runs: [{ text: 'Results' }],
		});
		const updated = updateTableOfContents(model, { newId })!;
		const entries = updated.blocks.slice(0, 2) as Paragraph[];
		expect(entries.map(text)).toEqual(['Lead Introduction', 'Results tail']);
		expect(at(entries, 0).runs.find((run) => run.fieldCode)?.fieldCode).toBe(' TOC \\o "1-1" ');
		expect(updated.blocks).toHaveLength(model.blocks.length + 1);
	});

	it('saves a TOC Word can update: field markers, tab stops and schema-ordered paragraph properties', async () => {
		const model = sample();
		model.blocks.unshift(
			...buildTableOfContents(model, { newId, pageNumbers: new Map([['h1', '1']]) }),
		);
		const { saveDocx } = await import('./index.js');
		const saved = await saveDocx(model);
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		expect(xml).toContain(
			'<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>',
		);
		expect(xml).toMatch(
			/<w:pPr><w:pStyle w:val="TOC1"\/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9360"\/><\/w:tabs><\/w:pPr>/,
		);
		expect(xml).toContain('<w:t>Introduction</w:t><w:tab/><w:t>1</w:t>');
		const reloaded = await loadDocx(saved);
		const first = reloaded.model.blocks[0] as Paragraph;
		expect(first.tabStops).toEqual([{ posTwips: 9360, align: 'right', leader: 'dot' }]);
		expect(findTableOfContents(reloaded.model.blocks)).toMatchObject({ start: 0, end: 1 });
	});

	it('links entries to _Toc bookmarks on the headings with PAGEREF page numbers', async () => {
		const model = sample();
		const entries = tocEntries(model);
		const { bookmarks, added } = tocBookmarks(model, entries);
		expect([...added.keys()]).toEqual(['h1', 'h2']);
		const toc = buildTableOfContents(model, {
			newId,
			bookmarks,
			pageNumbers: new Map([['h1', '1']]),
		});
		const name = bookmarks.get('h1')!;
		expect(at(toc, 0).runs.slice(3)).toEqual([
			{ text: 'Introduction\t', link: { anchor: name } },
			{ text: '', fieldChar: 'begin', link: { anchor: name } },
			{ text: '', fieldCode: ` PAGEREF ${name} \\h `, link: { anchor: name } },
			{ text: '', fieldChar: 'separate', link: { anchor: name } },
			{ text: '1', field: { instr: `PAGEREF ${name} \\h` }, link: { anchor: name } },
			{ text: '', fieldChar: 'end', link: { anchor: name } },
		]);
		const marked = withTocBookmarks(model, added);
		marked.blocks.unshift(...toc);
		const xml = await (
			await JSZip.loadAsync(await saveDocx(marked))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).toMatch(
			new RegExp(
				`<w:bookmarkStart w:id="(\\d+)" w:name="${name}"/><w:r><w:t>Introduction</w:t></w:r><w:bookmarkEnd w:id="\\1"/>`,
			),
		);
		expect(xml).toContain(`<w:hyperlink w:anchor="${name}" w:history="1">`);
		// Updating keeps the existing bookmarks rather than adding new ones.
		const reloaded = await loadDocx(await saveDocx(marked));
		const again = tocBookmarks(reloaded.model, tocEntries(reloaded.model));
		expect(again.added.size).toBe(0);
		expect(
			again.bookmarks.get(
				reloaded.model.blocks.find((b) => (b as Paragraph).style === 'Heading1')!.id,
			),
		).toBe(name);
	});
});
