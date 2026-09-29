import { twips } from '@christophervr/docx-core';
import { describe, expect, it } from 'vitest';
import { layoutSections } from './page-flow.js';
import { createFakeMeasurer } from './measure.js';
import type { LayoutDocumentInput, LayoutParagraph } from './input.js';

describe('layoutSections performance smoke test', () => {
	it('paginates 2,000 paragraphs in well under 2 seconds', () => {
		const blocks: LayoutParagraph[] = Array.from({ length: 2000 }, (_, i) => ({
			kind: 'paragraph',
			id: `p${i}`,
			runs: [
				{
					text: `Paragraph number ${i} with enough words to wrap across more than one line in a normal US Letter page column.`,
				},
			],
			spacingAfterTwips: twips(120),
		}));
		const input: LayoutDocumentInput = {
			sections: [
				{
					page: {
						widthPx: 816,
						heightPx: 1056,
						marginTopPx: 96,
						marginRightPx: 96,
						marginBottomPx: 96,
						marginLeftPx: 96,
					},
					blocks,
				},
			],
		};
		const measurer = createFakeMeasurer();
		const start = performance.now();
		const result = layoutSections(input, measurer);
		const elapsedMs = performance.now() - start;
		expect(result.pages.length).toBeGreaterThan(10);
		expect(elapsedMs).toBeLessThan(2000);
	});
});
