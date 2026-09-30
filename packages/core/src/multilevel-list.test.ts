import { describe, expect, it } from 'vitest';
import {
	computeListLabels,
	createDocument,
	ensureListDefinition,
	type Paragraph,
} from './index.js';

describe('multilevel list definitions', () => {
	it('numbers 1., 1.1. and 1.1.1. for the multilevel kind and 1. a) i. for outline', () => {
		for (const [kind, expected] of [
			['multilevel', ['1.', '1.1.', '1.1.1.', '2.']],
			['outline', ['1.', 'a)', 'i.', '2.']],
		] as const) {
			const model = createDocument();
			const { catalog, numId } = ensureListDefinition(model.numberingCatalog, kind);
			model.numberingCatalog = catalog;
			const levels = [0, 1, 2, 0];
			model.blocks = levels.map((level, i): Paragraph => ({
				type: 'paragraph',
				id: `p${i}`,
				runs: [{ text: 'x' }],
				numbering: { numId, level },
			}));
			const labels = computeListLabels(model);
			expect(model.blocks.map((b) => labels.get(b.id)?.text)).toEqual(expected);
		}
	});
});
