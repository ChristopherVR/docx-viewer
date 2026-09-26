import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx } from './index';
describe('Word core public package', () => {
	it('consumes the shared engine headlessly and preserves loaded package bytes', async () => {
		const model = createDocument();
		model.blocks = [
			{ type: 'paragraph', id: 'p1', runs: [{ text: 'Shared Word engine', bold: true }] },
		];
		const bytes = await saveDocx(model);
		const session = await loadDocx(bytes);
		expect(session.model.blocks[0]).toMatchObject({
			type: 'paragraph',
			runs: [{ text: 'Shared Word engine', bold: true }],
		});
		expect(await session.save()).toEqual(bytes);
	});
});
