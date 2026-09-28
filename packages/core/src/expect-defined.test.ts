import { describe, expect, it } from 'vitest';
import { expectDefined } from './expect-defined.js';

describe('expectDefined', () => {
	it('returns present values, including falsy ones', () => {
		expect(expectDefined(0, 'zero')).toBe(0);
		expect(expectDefined('', 'empty')).toBe('');
	});
	it('throws a descriptive error for missing values', () => {
		expect(() => expectDefined(undefined, 'table row 3')).toThrow(/table row 3 is missing/);
		expect(() => expectDefined(null, 'x')).toThrow(/x is missing/);
	});
});
