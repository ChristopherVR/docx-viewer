import { describe, expect, it } from 'vitest';
import { createDocument, type Paragraph } from 'docx-core';
import { modelToDoc, docToModel } from './model-adapter';
import { sameRuns } from './run-compare';
import { schema } from './schema';

const equation = {
	omml: '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
	display: false,
};

describe('equation model adaptation', () => {
	it('preserves source and surrounding run boundaries when another paragraph changes', () => {
		const model = createDocument();
		const runs = [{ text: 'before' }, { text: '', equation }, { text: 'after' }];
		model.blocks = [
			{ type: 'paragraph', id: 'math', runs },
			{ type: 'paragraph', id: 'edit', runs: [{ text: 'old' }] },
		];
		const doc = modelToDoc(model);
		expect(doc.firstChild!.child(1).type.name).toBe('equation');
		const changed = doc.copy(
			doc.content.replaceChild(
				1,
				schema.nodes.paragraph.create({ id: 'edit' }, schema.text('new')),
			),
		);
		const next = docToModel(changed, model);
		expect((next.blocks[0] as Paragraph).runs).toEqual(runs);
		expect(next.blocks[0]).toBe(model.blocks[0]);
	});

	it('does not merge equation atoms into text or treat changed source as unchanged', () => {
		const runs = [{ text: 'before' }, { text: '', equation }, { text: 'after' }];
		expect(sameRuns(runs, runs)).toBe(true);
		expect(sameRuns(runs, [{ text: 'beforeafter' }])).toBe(false);
		expect(
			sameRuns([{ text: '', equation }], [{ text: '', equation: { ...equation, display: true } }]),
		).toBe(false);
	});
});
