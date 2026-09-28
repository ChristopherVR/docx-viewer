import { describe, expect, it } from 'vitest';
import { ensureListDefinition } from './numbering-editing.js';
import { at, must } from './test-support/access.js';

describe('ensureListDefinition', () => {
	it('creates a fresh bulleted definition from an empty catalog', () => {
		const { catalog, numId } = ensureListDefinition(undefined, 'bullet');
		expect(numId).toBe(1);
		const num = must(catalog.nums['1'], 'numbering instance 1');
		const abstractNum = must(catalog.abstractNums[num.abstractNumId], 'abstract numbering');
		expect(at(abstractNum.levels, 0)).toMatchObject({ numFmt: 'bullet', lvlText: '•' });
		expect(Object.keys(abstractNum.levels)).toHaveLength(9);
	});

	it('creates a fresh decimal definition with cycling per-level formats', () => {
		const { catalog, numId } = ensureListDefinition(undefined, 'decimal');
		const num = must(catalog.nums[String(numId)], 'numbering instance');
		const abstractNum = must(catalog.abstractNums[num.abstractNumId], 'abstract numbering');
		expect(at(abstractNum.levels, 0)).toMatchObject({ numFmt: 'decimal', lvlText: '%1.' });
		expect(at(abstractNum.levels, 1)).toMatchObject({ numFmt: 'lowerLetter', lvlText: '%2.' });
		expect(at(abstractNum.levels, 2)).toMatchObject({ numFmt: 'lowerRoman', lvlText: '%3.' });
	});

	it('never reuses an existing numId or abstractNumId and preserves prior entries unchanged', () => {
		const first = ensureListDefinition(undefined, 'bullet');
		const second = ensureListDefinition(first.catalog, 'decimal');
		expect(second.numId).not.toBe(first.numId);
		expect(second.catalog.nums[String(first.numId)]).toEqual(
			first.catalog.nums[String(first.numId)],
		);
		expect(second.catalog.abstractNums).toMatchObject(first.catalog.abstractNums);
		// Independent restart: two decimal lists never share counters because they get separate numIds.
		const third = ensureListDefinition(second.catalog, 'decimal');
		expect(third.numId).not.toBe(second.numId);
	});

	it('numbers a numeric abstractNumId floor starting at 0 and numId floor starting at 1', () => {
		const { catalog } = ensureListDefinition(
			{
				abstractNums: { '3': { id: '3', levels: {} } },
				nums: { '7': { id: '7', abstractNumId: '3' } },
				warnings: [],
			},
			'bullet',
		);
		expect(Object.keys(catalog.abstractNums).sort()).toEqual(['3', '4']);
		expect(Object.keys(catalog.nums).sort()).toEqual(['7', '8']);
	});
});
