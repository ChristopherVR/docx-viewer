import { expect, it } from 'vitest';
import { createListDefinition, ensureListDefinition } from './numbering-editing.js';

it('copies custom levels and preserves existing independent list definitions', () => {
	const prior = ensureListDefinition(undefined, 'outline').catalog;
	const original = JSON.stringify(prior);
	const level = {
		level: 0,
		start: 3,
		numFmt: 'upperRoman' as const,
		lvlText: 'Part %1',
		suffix: 'space' as const,
	};
	const created = createListDefinition(prior, [level]);
	expect(created.numId).toBe(2);
	expect(JSON.stringify(prior)).toBe(original);
	level.start = 99;
	expect(created.catalog.abstractNums['1']!.levels[0]!.start).toBe(3);
	expect(created.catalog.abstractNums['0']).toBe(prior.abstractNums['0']);
});

it('rejects invalid definitions before they can enter an editable document', () => {
	const level = { level: 0, start: 1, numFmt: 'decimal' as const, lvlText: '%1.' };
	expect(() => createListDefinition(undefined, [])).toThrow('1 to 9');
	expect(() => createListDefinition(undefined, [level, level])).toThrow('distinct');
	expect(() => createListDefinition(undefined, [{ ...level, level: 9 }])).toThrow('0 through 8');
	expect(() => createListDefinition(undefined, [{ ...level, start: NaN }])).toThrow(
		'not valid WordprocessingML',
	);
	expect(() => createListDefinition(undefined, [{ ...level, numFmt: 'invalid' as never }])).toThrow(
		'not valid WordprocessingML',
	);
});
