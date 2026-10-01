import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './index.js';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('sections', () => {
	it('parses a paragraph-level section break followed by the final body section', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${ns}"><w:body>
<w:p><w:pPr><w:sectPr>
<w:type w:val="continuous"/>
<w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:num="2" w:space="360"/>
</w:sectPr></w:pPr><w:r><w:t>First section</w:t></w:r></w:p>
<w:p><w:r><w:t>Second section</w:t></w:r></w:p>
<w:sectPr>
<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
<w:pgNumType w:start="1" w:fmt="upperRoman"/>
<w:titlePg/>
<w:vAlign w:val="center"/>
<w:lnNumType/>
<w:pgBorders/>
</w:sectPr>
</w:body></w:document>`,
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		const sections = loaded.model.sections;
		expect(sections).toHaveLength(2);
		const [first, second] = sections!;
		expect(first).toMatchObject({
			endsAtBlockId: 'p0',
			type: 'continuous',
			pageWidthTwips: 12240,
			pageHeightTwips: 15840,
			orientation: 'portrait',
			headerDistanceTwips: 720,
			footerDistanceTwips: 720,
			gutterTwips: 0,
			columns: { count: 2, spacingTwips: 360, equalWidth: true },
		});
		expect(second).toMatchObject({
			endsAtBlockId: 'p1',
			type: 'nextPage',
			pageWidthTwips: 15840,
			pageHeightTwips: 12240,
			orientation: 'landscape',
			titlePage: true,
			verticalAlign: 'center',
			lineNumbering: true,
			pageBorders: {},
			pageNumbering: { start: 1, format: 'upperRoman' },
			columns: { count: 1, equalWidth: true },
		});
	});

	it('returns byte-exact source for a no-op save with sections present', async () => {
		const zip = new JSZip();
		zip.file(
			'word/document.xml',
			`<w:document xmlns:w="${ns}"><w:body><w:p><w:r><w:t>Only</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const loaded = await loadDocx(bytes);
		expect(loaded.model.sections).toHaveLength(1);
		expect(await loaded.save()).toEqual(bytes);
	});
});
