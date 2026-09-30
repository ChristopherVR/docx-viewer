import { describe, expect, it } from 'vitest';
import { isEastAsianChar, tokenizeRun } from './text-breaks.js';

describe('isEastAsianChar', () => {
	it('classifies CJK ideographs, kana, and hangul as East Asian', () => {
		expect(isEastAsianChar('中')).toBe(true); // 中
		expect(isEastAsianChar('あ')).toBe(true); // あ
		expect(isEastAsianChar('가')).toBe(true); // 가
		expect(isEastAsianChar('a')).toBe(false);
		expect(isEastAsianChar(' ')).toBe(false);
	});
});

describe('tokenizeRun', () => {
	it('does not break a CJK grapheme from its combining mark', () => {
		expect(tokenizeRun('あ\u3099中', 0)).toEqual([
			{ kind: 'word', text: 'あ\u3099', runIndex: 0, sourceStart: 0 },
			{ kind: 'word', text: '中', runIndex: 0, sourceStart: 2 },
		]);
	});
	it('splits ordinary words at spaces without breaking inside a word', () => {
		const tokens = tokenizeRun('hello world', 0);
		expect(tokens.map((t) => t.kind)).toEqual(['word', 'space', 'word']);
		expect(tokens[0]).toMatchObject({ text: 'hello', sourceStart: 0 });
		expect(tokens[2]).toMatchObject({ text: 'world', sourceStart: 6 });
	});

	it('allows a break after a hyphen but not after a non-breaking hyphen', () => {
		const hyphenated = tokenizeRun('well-known', 0);
		expect(hyphenated.filter((t) => t.kind === 'word').map((t) => t.text)).toEqual([
			'well-',
			'known',
		]);
		const nonBreaking = tokenizeRun('well‑known', 0);
		expect(nonBreaking.filter((t) => t.kind === 'word').map((t) => t.text)).toEqual(['well‑known']);
	});

	it('treats each CJK character as its own breakable unit', () => {
		const tokens = tokenizeRun('中文字', 0);
		expect(tokens.map((t) => t.kind)).toEqual(['word', 'word', 'word']);
		expect(tokens.map((t) => (t as { text: string }).text)).toEqual(['中', '文', '字']);
	});

	it('emits tab and line-break tokens', () => {
		const tokens = tokenizeRun('a\tb\nc', 0);
		expect(tokens.map((t) => t.kind)).toEqual(['word', 'tab', 'word', 'lineBreak', 'word']);
	});
});
