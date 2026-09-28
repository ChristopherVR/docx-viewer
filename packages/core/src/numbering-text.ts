// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word list-format numeral conversions (WordprocessingML ST_NumberFormat).

/**
 * Roman numeral conversion. Adapted from
 * pptx-viewer-new packages/core/src/core/utils/auto-number-format.ts (romanNumeral),
 * which documents the same clamp-to-[1,3999] behavior verified against real output.
 */
export function romanNumeral(n: number): string {
	const values: readonly number[] = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
	const numerals: readonly string[] = [
		'M',
		'CM',
		'D',
		'CD',
		'C',
		'XC',
		'L',
		'XL',
		'X',
		'IX',
		'V',
		'IV',
		'I',
	];
	let remaining = Math.max(1, Math.min(Math.trunc(n), 3999));
	let result = '';
	for (const [i, value] of values.entries()) {
		while (remaining >= value) {
			result += numerals[i] ?? '';
			remaining -= value;
		}
	}
	return result;
}

const LATIN_LOWER = Array.from('abcdefghijklmnopqrstuvwxyz');
/**
 * Word's alphabetic label (lower-case): past `z` the letter repeats rather than
 * counting spreadsheet-style. Adapted from the `repeatedLabel` rule in
 * pptx-viewer-new packages/core/src/core/utils/auto-number-alphabets.ts
 * (1 -> "a", 26 -> "z", 27 -> "aa", 52 -> "zz", 53 -> "aaa").
 */
export function letterLabel(n: number): string {
	const index = Math.max(0, Math.trunc(n) - 1);
	const letter = LATIN_LOWER[index % LATIN_LOWER.length] ?? 'a';
	return letter.repeat(Math.floor(index / LATIN_LOWER.length) + 1);
}

const ORDINAL_SUFFIX = (n: number): string => {
	const abs = Math.abs(Math.trunc(n));
	if (abs % 100 >= 11 && abs % 100 <= 13) return 'th';
	switch (abs % 10) {
		case 1:
			return 'st';
		case 2:
			return 'nd';
		case 3:
			return 'rd';
		default:
			return 'th';
	}
};
/** `decimal` numeral immediately followed by its English ordinal suffix, e.g. `2nd`. */
export function ordinalNumeral(n: number): string {
	return `${Math.trunc(n)}${ORDINAL_SUFFIX(n)}`;
}

const ONES = [
	'',
	'one',
	'two',
	'three',
	'four',
	'five',
	'six',
	'seven',
	'eight',
	'nine',
	'ten',
	'eleven',
	'twelve',
	'thirteen',
	'fourteen',
	'fifteen',
	'sixteen',
	'seventeen',
	'eighteen',
	'nineteen',
];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
/** English cardinal words for 0..999,999 (falls back to the decimal numeral outside that range). */
export function cardinalWords(n: number): string {
	const value = Math.trunc(n);
	if (value < 0 || value > 999999) return String(value);
	if (value === 0) return 'zero';
	const chunk = (v: number): string => {
		const parts: string[] = [];
		if (v >= 100) {
			parts.push(ONES[Math.floor(v / 100)] ?? '', 'hundred');
			v %= 100;
		}
		if (v >= 20) {
			parts.push((TENS[Math.floor(v / 10)] ?? '') + (v % 10 ? `-${ONES[v % 10] ?? ''}` : ''));
		} else if (v > 0) {
			parts.push(ONES[v] ?? '');
		}
		return parts.join(' ');
	};
	const thousands = Math.floor(value / 1000);
	const rest = value % 1000;
	const words: string[] = [];
	if (thousands) words.push(chunk(thousands), 'thousand');
	if (rest || !thousands) words.push(chunk(rest));
	return words.filter(Boolean).join(' ');
}

const ORDINAL_WORDS: Record<string, string> = {
	one: 'first',
	two: 'second',
	three: 'third',
	five: 'fifth',
	eight: 'eighth',
	nine: 'ninth',
	twelve: 'twelfth',
};
/** English ordinal words for 0..999,999, e.g. `twenty-first` (falls back to the ordinal numeral). */
export function ordinalWords(n: number): string {
	const value = Math.trunc(n);
	if (value < 0 || value > 999999) return ordinalNumeral(value);
	const words = cardinalWords(value);
	const splitAt = Math.max(words.lastIndexOf(' '), words.lastIndexOf('-'));
	const head = splitAt >= 0 ? words.slice(0, splitAt + 1) : '';
	const lastWord = splitAt >= 0 ? words.slice(splitAt + 1) : words;
	if (ORDINAL_WORDS[lastWord]) return head + ORDINAL_WORDS[lastWord];
	if (lastWord.endsWith('y')) return `${head}${lastWord.slice(0, -1)}ieth`;
	return `${head}${lastWord}th`;
}
