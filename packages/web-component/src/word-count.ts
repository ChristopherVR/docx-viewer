/** Unicode-aware word counting; segmentation is provided by the host's ICU data. */
export function countWords(text: string, locale?: string): number {
	if (typeof Intl.Segmenter === 'function') {
		let segmenter: Intl.Segmenter;
		try {
			segmenter = new Intl.Segmenter(locale || undefined, { granularity: 'word' });
		} catch {
			segmenter = new Intl.Segmenter(undefined, { granularity: 'word' });
		}
		return [...segmenter.segment(text)].filter((part) => part.isWordLike).length;
	}
	// Older engines cannot provide dictionary segmentation for unspaced scripts.
	return text.match(/[\p{L}\p{N}][\p{L}\p{M}\p{N}'’]*/gu)?.length ?? 0;
}
