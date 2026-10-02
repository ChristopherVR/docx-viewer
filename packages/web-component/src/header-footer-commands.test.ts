import { createDocument, type DocumentModel, type Paragraph } from 'docx-core';
import { describe, expect, it } from 'vitest';
import {
	nextPartName,
	pageNumberParagraph,
	withBlankHeaderFooter,
	withPageNumber,
} from './header-footer-commands';

let n = 0;
const id = () => `t${n++}`;

function withBody(): DocumentModel {
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Body' }] }];
	return model;
}
const footerParagraphs = (model: DocumentModel) =>
	(model.sections![0]!.footers!.default!.blocks as Paragraph[]) ?? [];

describe('withPageNumber', () => {
	it('adds numbers to the inherited story without creating a local reference', () => {
		const model = withBlankHeaderFooter(withBody(), 'headers', id);
		const section = model.sections![0]!;
		model.sections = [section, { ...section, endsAtBlockId: 'second', headers: {} }];
		const next = withPageNumber(model, 'top', 'center', id, 1, 'pageOfTotal');
		expect(next.sections![1]!.headers!.default).toBeUndefined();
		expect((next.sections![0]!.headers!.default!.blocks[1] as Paragraph).runs).toHaveLength(4);
	});
	it('switches generated page patterns without duplicates and keeps surrounding paragraphs', () => {
		const model = withBlankHeaderFooter(withBody(), 'footers', id);
		footerParagraphs(model)[0]!.runs = [{ text: 'Confidential' }];
		const numbered = withPageNumber(model, 'bottom', 'center', id, 0, 'pageOfTotal');
		const page = footerParagraphs(numbered)[1]!;
		expect(page.runs.map((run) => run.field?.instr.trim() ?? run.text)).toEqual([
			'Page ',
			'PAGE',
			' of ',
			'NUMPAGES',
		]);
		const again = withPageNumber(numbered, 'bottom', 'right', id, 0, 'pageOfTotal');
		expect(footerParagraphs(again)).toHaveLength(2);
		expect(footerParagraphs(again)[0]!.runs[0]!.text).toBe('Confidential');
		const plain = withPageNumber(again, 'bottom', 'left', id, 0, 'plain');
		expect(footerParagraphs(plain)[1]!.runs).toEqual([
			{ text: '1', field: { instr: ' PAGE ', simple: true } },
		]);
		expect(page.runs).toHaveLength(4);
	});
	it('keeps imported custom page text and complex fields when adding the total', () => {
		const model = withPageNumber(withBody(), 'bottom', 'left', id);
		const runs = [
			{ text: 'Company, page ' },
			{ text: '', fieldChar: 'begin' as const },
			{ text: '', fieldCode: ' PAGE ' },
			{ text: '', fieldChar: 'separate' as const },
			{ text: '4' },
			{ text: '', fieldChar: 'end' as const },
		];
		footerParagraphs(model)[0]!.runs = runs;
		const changed = withPageNumber(model, 'bottom', 'center', id, 0, 'pageOfTotal');
		expect(footerParagraphs(changed)[0]!.runs.slice(0, 6)).toEqual(runs);
		expect(footerParagraphs(changed)[0]!.runs[7]!.field?.instr.trim()).toBe('NUMPAGES');
		expect(
			footerParagraphs(withPageNumber(changed, 'bottom', 'center', id, 0, 'pageOfTotal'))[0]!.runs,
		).toHaveLength(8);
	});
	it('targets a later section and updates every slot linked to its existing part', () => {
		const model = withBody();
		const section = withPageNumber(model, 'bottom', 'left', id).sections![0]!;
		const { footers: _footer, ...noFooter } = section;
		model.sections = [
			{ ...noFooter, endsAtBlockId: 'before' },
			{ ...section, endsAtBlockId: 'p1' },
		];
		const changed = withPageNumber(model, 'bottom', 'right', id, 1);
		expect(changed.sections![0]!.footers).toBeUndefined();
		expect((changed.sections![1]!.footers!.default!.blocks[0] as Paragraph).align).toBe('right');
		model.sections = [section, { ...section, endsAtBlockId: 'other' }];
		const linked = withPageNumber(model, 'bottom', 'center', id, 1);
		expect(
			linked.sections!.map((part) => (part.footers!.default!.blocks[0] as Paragraph).align),
		).toEqual(['center', 'center']);
	});
	it('creates a footer with a PAGE field for a document that has none', () => {
		const next = withPageNumber(withBody(), 'bottom', 'center', id);
		const footer = next.sections![0]!.footers!.default!;
		expect(footer.partName).toBe('word/footer1.xml');
		const paragraph = footer.blocks[0] as Paragraph;
		expect(paragraph.align).toBe('center');
		expect(paragraph.runs[0]?.field).toEqual({ instr: ' PAGE ', simple: true });
		expect(next.sections![0]!.headers).toBeUndefined();
	});

	it('uses the header for top positions', () => {
		const next = withPageNumber(withBody(), 'top', 'right', id);
		expect(next.sections![0]!.headers!.default!.partName).toBe('word/header1.xml');
		expect(next.sections![0]!.footers).toBeUndefined();
	});

	it('does not mutate the model it was given', () => {
		const model = withBody();
		withPageNumber(model, 'bottom', 'left', id);
		expect(model.sections).toBeUndefined();
	});

	it('re-aligns an existing page-number paragraph instead of adding another', () => {
		const once = withPageNumber(withBody(), 'bottom', 'left', id);
		const twice = withPageNumber(once, 'bottom', 'right', id);
		expect(footerParagraphs(twice)).toHaveLength(1);
		expect(footerParagraphs(twice)[0]!.align).toBe('right');
		expect(footerParagraphs(once)[0]!.align).toBe('left');
	});

	it('keeps existing footer text and appends the page number', () => {
		const model = withBody();
		model.sections = [
			{
				...withPageNumber(withBody(), 'bottom', 'left', id).sections![0]!,
				footers: {
					default: {
						partName: 'word/footer3.xml',
						blocks: [{ type: 'paragraph', id: 'f', runs: [{ text: 'Company confidential' }] }],
					},
				},
			},
		];
		const next = withPageNumber(model, 'bottom', 'center', id);
		const blocks = footerParagraphs(next);
		expect(blocks).toHaveLength(2);
		expect(blocks[0]!.runs[0]!.text).toBe('Company confidential');
		expect(blocks[1]!.runs[0]!.field?.instr.trim()).toBe('PAGE');
		expect(next.sections![0]!.footers!.default!.partName).toBe('word/footer3.xml');
	});

	it('recognises a complex PAGE field written with field codes', () => {
		const model = withBody();
		const base = withPageNumber(model, 'bottom', 'left', id);
		base.sections![0]!.footers!.default!.blocks = [
			{
				type: 'paragraph',
				id: 'c',
				runs: [
					{ text: '', fieldChar: 'begin' },
					{ text: '', fieldCode: ' PAGE ' },
					{ text: '', fieldChar: 'separate' },
					{ text: '4' },
					{ text: '', fieldChar: 'end' },
				],
			},
		];
		const next = withPageNumber(base, 'bottom', 'right', id);
		expect(footerParagraphs(next)).toHaveLength(1);
		expect(footerParagraphs(next)[0]!.align).toBe('right');
	});
});

describe('withBlankHeaderFooter', () => {
	it('creates separate first/even parts and enables only their required settings', () => {
		const normal = withBlankHeaderFooter(withBody(), 'footers', id);
		const first = withBlankHeaderFooter(normal, 'footers', id, 0, 'first');
		expect(first.sections![0]!.titlePage).toBe(true);
		expect(first.evenAndOddHeaders).toBeUndefined();
		expect(first.sections![0]!.footers!.default).toEqual(normal.sections![0]!.footers!.default);
		expect(first.sections![0]!.footers!.first!.partName).toBe('word/footer2.xml');
		const even = withPageNumber(first, 'bottom', 'right', id, 0, 'pageOfTotal', 'even');
		expect(even.evenAndOddHeaders).toBe(true);
		expect(even.sections![0]!.footers!.even!.partName).toBe('word/footer3.xml');
		expect(even.sections![0]!.footers!.first).toEqual(first.sections![0]!.footers!.first);
		expect(withBlankHeaderFooter(even, 'footers', id, 0, 'even')).toBe(even);
	});
	it('reactivates existing first-page content without replacing it', () => {
		const first = withBlankHeaderFooter(withBody(), 'headers', id, 0, 'first');
		first.sections![0]!.titlePage = false;
		(first.sections![0]!.headers!.first!.blocks[0] as Paragraph).runs = [{ text: 'Letterhead' }];
		const enabled = withBlankHeaderFooter(first, 'headers', id, 0, 'first');
		expect(enabled.sections![0]!.titlePage).toBe(true);
		expect(enabled.sections![0]!.headers!.first).toEqual(first.sections![0]!.headers!.first);
	});
	it('creates an empty part once and is a no-op afterwards', () => {
		const model = withBody();
		const created = withBlankHeaderFooter(model, 'headers', id);
		expect(created.sections![0]!.headers!.default!.blocks).toHaveLength(1);
		expect(withBlankHeaderFooter(created, 'headers', id)).toBe(created);
	});
});

describe('nextPartName', () => {
	it('numbers past the highest part the model already uses', () => {
		const model = withBody();
		expect(nextPartName(model, 'footers')).toBe('word/footer1.xml');
		const withTwo = withPageNumber(withBody(), 'bottom', 'left', id);
		withTwo.sections![0]!.footers!.default!.partName = 'word/footer7.xml';
		expect(nextPartName(withTwo, 'footers')).toBe('word/footer8.xml');
		expect(nextPartName(withTwo, 'headers')).toBe('word/header1.xml');
	});
});

describe('pageNumberParagraph', () => {
	it('shows 1 as the placeholder result of the field', () => {
		expect(pageNumberParagraph('x', 'left').runs[0]!.text).toBe('1');
	});
});

describe('rebuilding the model keeps created headers and footers', () => {
	it('restores the absence of a footer from an earlier document snapshot', async () => {
		const { docToModel, modelToDoc } = await import('./model-adapter');
		const withFooter = withPageNumber(withBody(), 'bottom', 'right', id);
		// A new document has no `sections` attribute on the editor document.
		const { sections: _kept, ...withoutSections } = withFooter;
		const doc = modelToDoc(withoutSections);
		expect(doc.attrs.sections).toBeNull();
		const rebuilt = docToModel(doc, withFooter);
		expect(rebuilt.sections).toHaveLength(1);
		expect(rebuilt.sections![0]!.footers).toBeUndefined();
		expect(rebuilt.sections![0]!.pageWidthTwips).toBeGreaterThan(0);
	});

	it('does not invent sections for a document without header or footer content', async () => {
		const { docToModel, modelToDoc } = await import('./model-adapter');
		const plain = withBody();
		const rebuilt = docToModel(modelToDoc(plain), plain);
		expect(rebuilt.sections).toBeUndefined();
	});

	it('uses the recorded layout, with content by section, when there is one', async () => {
		const { docToModel, modelToDoc } = await import('./model-adapter');
		const withFooter = withPageNumber(withBody(), 'bottom', 'left', id);
		const doc = modelToDoc(withFooter);
		expect(typeof doc.attrs.sections).toBe('string');
		expect(docToModel(doc, withFooter).sections![0]!.footers!.default!.partName).toBe(
			'word/footer1.xml',
		);
	});
});
