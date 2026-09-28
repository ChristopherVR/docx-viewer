import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { saveDocx } from './save.js';
import { createDocument } from './index.js';
import { acceptRevision, linkedRevisionIds, rejectRevision } from './revision-commands.js';
import type { DocumentModel, Paragraph } from './model.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const A = 'w:author="Ada" w:date="2024-01-01T00:00:00Z"';

// "Moved one" + "Moved two" move from the first two paragraphs to the end of the last one.
const BODY =
	`<w:p><w:moveFromRangeStart w:id="10" ${A} w:name="move1"/><w:moveFrom w:id="11" ${A}><w:r><w:t>Moved one</w:t></w:r></w:moveFrom></w:p>` +
	`<w:p><w:moveFrom w:id="12" ${A}><w:r><w:t>Moved two</w:t></w:r></w:moveFrom><w:moveFromRangeEnd w:id="10"/><w:r><w:t xml:space="preserve"> stays</w:t></w:r></w:p>` +
	`<w:p><w:r><w:t xml:space="preserve">End </w:t></w:r><w:moveToRangeStart w:id="13" ${A} w:name="move1"/><w:moveTo w:id="14" ${A}><w:r><w:t>Moved one two</w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="13"/></w:p>`;

async function load() {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document ${NS}><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const xmlOf = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const texts = (model: DocumentModel) =>
	model.blocks.map((block) => (block as Paragraph).runs.map((run) => run.text).join(''));

describe('tracked moves', () => {
	it('links both sides of a move by name across paragraphs', async () => {
		const { model } = await load();
		const runs = model.blocks.flatMap((block) => (block as Paragraph).runs);
		expect(runs.filter((run) => run.revision).map((run) => run.revision)).toEqual([
			expect.objectContaining({
				kind: 'moveFrom',
				id: '11',
				move: { name: 'move1', rangeId: '10' },
			}),
			expect.objectContaining({
				kind: 'moveFrom',
				id: '12',
				move: { name: 'move1', rangeId: '10' },
			}),
			expect.objectContaining({ kind: 'moveTo', id: '14', move: { name: 'move1', rangeId: '13' } }),
		]);
		expect(linkedRevisionIds(model, '14').sort()).toEqual(['11', '12', '14']);
	});

	it('accepting either side completes the move; rejecting restores it', async () => {
		const { model } = await load();
		expect(texts(acceptRevision(model, '12'))).toEqual(['', ' stays', 'End Moved one two']);
		expect(texts(rejectRevision(model, '14'))).toEqual(['Moved one', 'Moved two stays', 'End ']);
		const rejected = rejectRevision(model, '14');
		expect(
			rejected.blocks.flatMap((block) => (block as Paragraph).runs).some((run) => run.revision),
		).toBe(false);
	});

	it('saves edited paragraphs with one move range per side, keeping range ids', async () => {
		const loaded = await load();
		const blocks = loaded.model.blocks.map((block) => ({
			...(block as Paragraph),
			runs: (block as Paragraph).runs.map((run) =>
				run.revision ? run : { ...run, text: run.text.toUpperCase() },
			),
		}));
		const xml = await xmlOf(await loaded.save({ ...loaded.model, blocks }));
		expect(xml.match(/<w:moveFromRangeStart /g)).toHaveLength(1);
		expect(xml.match(/<w:moveFromRangeEnd /g)).toHaveLength(1);
		expect(xml).toMatch(
			/<w:moveFromRangeStart w:id="10" w:author="Ada" w:date="[^"]+" w:name="move1"\/><w:moveFrom w:id="11"/,
		);
		expect(xml).toMatch(
			/Moved two<\/w:delText><\/w:r><\/w:moveFrom><w:moveFromRangeEnd w:id="10"\/><w:r><w:t xml:space="preserve"> STAYS/,
		);
		expect(xml).toMatch(
			/END <\/w:t><\/w:r><w:moveToRangeStart w:id="13"[^>]*w:name="move1"\/><w:moveTo w:id="14"/,
		);
		const reloaded = await loadDocx(await loaded.save({ ...loaded.model, blocks }));
		expect(linkedRevisionIds(reloaded.model, '11').sort()).toEqual(['11', '12', '14']);
	});

	it('numbers range markers for moves created in the editor after existing ids', async () => {
		const model = createDocument();
		const revision = { author: 'Ada', id: '5', move: { name: 'move9' } };
		model.blocks = [
			{
				type: 'paragraph',
				id: 'a',
				runs: [{ text: 'Go', revision: { ...revision, kind: 'moveFrom' } }],
			},
			{
				type: 'paragraph',
				id: 'b',
				runs: [{ text: 'Go', revision: { ...revision, id: '6', kind: 'moveTo' } }],
			},
		];
		const xml = await xmlOf(await saveDocx(model));
		const starts = [...xml.matchAll(/<w:(moveFrom|moveTo)RangeStart w:id="(\d+)"/g)].map((m) => [
			m[1],
			m[2],
		]);
		expect(starts).toEqual([
			['moveFrom', '7'],
			['moveTo', '8'],
		]);
		expect(xml).toContain('<w:moveFromRangeEnd w:id="7"/>');
		expect(xml).toContain('<w:moveToRangeEnd w:id="8"/>');
	});
});

describe('resolved revisions leave no marker key behind', () => {
	it('drops the revision key instead of writing undefined, so JSON and key checks stay clean', async () => {
		const { model } = await load();
		const accepted = acceptRevision(model, '14');
		for (const block of accepted.blocks)
			for (const run of (block as Paragraph).runs) expect('revision' in run).toBe(false);
		expect(JSON.parse(JSON.stringify(accepted))).toEqual(accepted);
	});
});
