import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Paragraph } from './model.js';

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

describe('revision ids', () => {
	it('renumbers editor revision ids to unused decimals, consistently across paragraphs', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document ${NS}><w:body><w:p><w:bookmarkStart w:id="7" w:name="b"/><w:r><w:t>One</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p><w:p><w:ins w:id="3" w:author="A"><w:r><w:t>Two</w:t></w:r></w:ins></w:p><w:sectPr/></w:body></w:document>`,
		);
		const loaded = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
		const revision = { kind: 'insert' as const, author: 'Ada', id: 'dve-rev-abc-1' };
		const [first, second] = loaded.model.blocks as Paragraph[];
		const blocks = [
			{ ...first, runs: [...first.runs, { text: ' new', revision }] },
			{
				...second,
				runs: [...second.runs, { text: ' more', revision }],
				markRevision: { ...revision, id: 'dve-rev-abc-2' },
			},
		];
		const xml = await (
			await JSZip.loadAsync(await loaded.save({ ...loaded.model, blocks }))
		)
			.file('word/document.xml')!
			.async('string');
		expect(xml).not.toContain('dve-rev');
		// Existing ids (3, 7) are kept; new ones start after the largest id in the part.
		expect(xml.match(/<w:ins w:id="8" /g)).toHaveLength(2);
		expect(xml).toContain('<w:ins w:id="3" ');
		expect(xml).toMatch(/<w:rPr><w:ins w:id="9" /);
	});
});
