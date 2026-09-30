import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import { computeListLabels } from './numbering-format.js';
import { restartCases, restartFixture } from './test-support/restart-fixture.js';

for (const { name, restart, expected } of restartCases) {
	it(`renders Word's ${name} restart rule and preserves source definitions during text editing`, async () => {
		const bytes = await restartFixture(restart);
		const loaded = await loadDocx(bytes);
		const labels = computeListLabels(loaded.model);
		expect(loaded.model.blocks.map((p) => labels.get(p.id)?.text)).toEqual(
			expected.map((n) => `${n}.`),
		);
		const paragraph = loaded.model.blocks[0];
		if (paragraph?.type !== 'paragraph') throw new Error('Missing first paragraph');
		paragraph.runs[0]!.text = 'Edited';
		const output = await JSZip.loadAsync(await loaded.save(loaded.model));
		const input = await JSZip.loadAsync(bytes);
		expect(await output.file('word/numbering.xml')!.async('string')).toBe(
			await input.file('word/numbering.xml')!.async('string'),
		);
		const reloaded = await loadDocx(await output.generateAsync({ type: 'uint8array' }));
		const again = computeListLabels(reloaded.model);
		expect(reloaded.model.blocks.map((p) => again.get(p.id)?.text)).toEqual(
			expected.map((n) => `${n}.`),
		);
	});
}

it('restarts a skipped level independently of an intermediate never-restart level', async () => {
	const { model } = await loadDocx(await restartFixture(undefined, true));
	const labels = computeListLabels(model);
	expect(model.blocks.map((p) => labels.get(p.id)?.text)).toEqual([
		'1.',
		'1.',
		'1.',
		'2.',
		'2.',
		'1.',
		'2.',
		'1.',
	]);
});

it('resolves restart and start overrides per numbering instance without sharing counters', async () => {
	const { model } = await loadDocx(await restartFixture(undefined));
	const catalog = model.numberingCatalog!;
	const level = catalog.abstractNums['0']!.levels[2]!;
	catalog.nums['2'] = {
		id: '2',
		abstractNumId: '0',
		levelOverrides: { 2: { startOverride: 7, lvl: { ...level, lvlRestart: 0 } } },
	};
	model.blocks = [2, 1, 2, 1, 2].map((numId, i) => ({
		type: 'paragraph',
		id: `p${i}`,
		runs: [{ text: 'x' }],
		numbering: { numId, level: i === 3 ? 0 : 2 },
	}));
	const labels = computeListLabels(model);
	expect(model.blocks.map((p) => labels.get(p.id)?.text)).toEqual(['7.', '1.', '8.', '2.', '9.']);
});

it('initializes omitted ancestors when an outline starts at a nested level', async () => {
	const { model } = await loadDocx(await restartFixture(undefined));
	model.blocks = [2, 0, 1, 2].map((level, i) => ({
		type: 'paragraph',
		id: `p${i}`,
		runs: [{ text: 'x' }],
		numbering: { numId: 1, level },
	}));
	const labels = computeListLabels(model);
	expect(model.blocks.map((p) => labels.get(p.id)?.text)).toEqual(['1.', '2.', '1.', '1.']);
});
