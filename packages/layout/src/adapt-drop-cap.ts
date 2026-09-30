import type { Block } from '@christophervr/docx-core';
import type { LayoutBlock } from './input.js';

/**
 * Print Layout has no text frames, so a drop cap paragraph (`w:framePr w:dropCap`) is folded into
 * the paragraph after it as an enlarged initial letter: a raised cap, not Word's dropped one.
 * `blocks` stays aligned with `source`; the returned set holds the folded (dropped) blocks so
 * section slices can leave them out.
 */
export function foldDropCaps(source: Block[], blocks: LayoutBlock[]): Set<LayoutBlock> {
	const folded = new Set<LayoutBlock>();
	source.forEach((block, index) => {
		if (block.type !== 'paragraph' || !block.dropCap) return;
		const cap = blocks[index];
		const next = blocks[index + 1];
		if (cap?.kind !== 'paragraph' || next?.kind !== 'paragraph') return;
		next.runs = [...cap.runs, ...next.runs];
		folded.add(cap);
	});
	return folded;
}
