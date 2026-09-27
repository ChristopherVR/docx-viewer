import type { MarkSpec } from 'prosemirror-model';

/** Stable, accessible palette for attributing tracked-change authors by name. */
const REVIEW_AUTHOR_PALETTE = [
	'#c2431f',
	'#1a73e8',
	'#188038',
	'#8430ce',
	'#b2103c',
	'#00838f',
	'#e37400',
	'#5b6bc0',
];
export function authorColor(author: string): string {
	let hash = 0;
	for (let i = 0; i < author.length; i++) hash = (hash * 31 + author.charCodeAt(i)) >>> 0;
	return REVIEW_AUTHOR_PALETTE[hash % REVIEW_AUTHOR_PALETTE.length];
}

/** Marks for tracked-change insertions/deletions and comment-range anchors; merged into schema.ts. */
export const reviewMarks: Record<string, MarkSpec> = {
	insertion: {
		attrs: { author: { default: '' }, date: { default: null }, id: { default: '' } },
		parseDOM: [
			{
				tag: 'ins[data-revision-id]',
				getAttrs: (el) => ({
					author: (el as HTMLElement).dataset.author || '',
					date: (el as HTMLElement).dataset.date || null,
					id: (el as HTMLElement).dataset.revisionId || '',
				}),
			},
		],
		toDOM: (mark) => [
			'ins',
			{
				'data-revision-id': mark.attrs.id,
				'data-author': mark.attrs.author,
				...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
				class: 'dve-revision-insert',
				style: `--dve-revision-color:${authorColor(mark.attrs.author)}`,
			},
			0,
		],
	},
	deletion: {
		attrs: { author: { default: '' }, date: { default: null }, id: { default: '' } },
		parseDOM: [
			{
				tag: 'del[data-revision-id]',
				getAttrs: (el) => ({
					author: (el as HTMLElement).dataset.author || '',
					date: (el as HTMLElement).dataset.date || null,
					id: (el as HTMLElement).dataset.revisionId || '',
				}),
			},
		],
		toDOM: (mark) => [
			'del',
			{
				'data-revision-id': mark.attrs.id,
				'data-author': mark.attrs.author,
				...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
				class: 'dve-revision-delete',
				style: `--dve-revision-color:${authorColor(mark.attrs.author)}`,
			},
			0,
		],
	},
	comment: {
		attrs: { ids: { default: [] } },
		parseDOM: [
			{
				tag: 'span[data-comment-ids]',
				getAttrs: (el) => ({
					ids: ((el as HTMLElement).dataset.commentIds || '').split(',').filter(Boolean),
				}),
			},
		],
		toDOM: (mark) => [
			'span',
			{ class: 'dve-comment-range', 'data-comment-ids': (mark.attrs.ids as string[]).join(',') },
			0,
		],
	},
};
