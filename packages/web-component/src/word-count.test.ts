import { describe, expect, it } from 'vitest';
import { countWords } from './word-count';

describe('Unicode word counts', () => {
	it('counts words rather than punctuation or emoji', () => {
		expect(countWords('Hello, world! 👨‍👩‍👧‍👦 — café cafe\u0301', 'en')).toBe(4);
		expect(countWords('  … 👩🏽‍💻\n', 'en')).toBe(0);
	});
	it('segments Arabic, Hebrew and unspaced Japanese', () => {
		expect(countWords('مرحبا بالعالم', 'ar')).toBe(2);
		expect(countWords('שלום עולם', 'he')).toBe(2);
		expect(countWords('これは日本語です', 'ja')).toBeGreaterThan(1);
	});
	it('falls back to the runtime locale for an invalid host language tag', () => {
		expect(countWords('two words', 'not_a_tag')).toBe(2);
	});
});
