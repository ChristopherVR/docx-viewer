import type { NodeSpec, MarkSpec } from 'prosemirror-model';

const safeAttrValue = (value: unknown): string => String(value ?? '').replace(/[;{}"]/g, '');

/** Inline picture atom: a placeholder span for unsupported drawings, an `<img>` for real pictures.
 * Actual `src` resolution from package media bytes happens in a node view (image-media.ts), not here. */
export const imageNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: true,
	attrs: {
		relId: { default: '' },
		partName: { default: '' },
		contentType: { default: '' },
		widthPx: { default: 96 },
		heightPx: { default: 96 },
		altText: { default: null },
		title: { default: null },
		anchored: { default: false },
		unsupported: { default: null },
	},
	parseDOM: [
		{
			tag: 'img[data-docx-image], span[data-docx-image-placeholder]',
			getAttrs: (el) => ({
				relId: (el as HTMLElement).dataset.relId || '',
				partName: (el as HTMLElement).dataset.partName || '',
				contentType: (el as HTMLElement).dataset.contentType || '',
				widthPx: Number((el as HTMLElement).dataset.widthPx) || 96,
				heightPx: Number((el as HTMLElement).dataset.heightPx) || 96,
				altText: (el as HTMLElement).getAttribute('alt') || null,
				title: (el as HTMLElement).getAttribute('title') || null,
				anchored: (el as HTMLElement).dataset.anchored === '1',
				unsupported: (el as HTMLElement).dataset.unsupported || null,
			}),
		},
	],
	toDOM: (node) =>
		node.attrs.unsupported
			? [
					'span',
					{
						'data-docx-image-placeholder': '1',
						class: 'dve-image-placeholder',
						style: `width:${node.attrs.widthPx}px;height:${node.attrs.heightPx}px`,
					},
					node.attrs.unsupported,
				]
			: [
					'img',
					{
						'data-docx-image': '1',
						'data-rel-id': node.attrs.relId,
						'data-part-name': node.attrs.partName,
						'data-content-type': node.attrs.contentType,
						'data-width-px': String(node.attrs.widthPx),
						'data-height-px': String(node.attrs.heightPx),
						'data-anchored': node.attrs.anchored ? '1' : '0',
						class: node.attrs.anchored ? 'dve-image dve-image-anchored' : 'dve-image',
						alt: node.attrs.altText || '',
						title: node.attrs.title || '',
						width: String(node.attrs.widthPx),
						height: String(node.attrs.heightPx),
					},
				],
};

/** `w:hyperlink` (external `href` or internal `anchor`), with an optional tooltip. */
export const linkMarkSpec: MarkSpec = {
	attrs: {
		href: { default: null },
		anchor: { default: null },
		tooltip: { default: null },
	},
	inclusive: false,
	parseDOM: [
		{
			tag: 'a[data-docx-link]',
			getAttrs: (el) => ({
				href: (el as HTMLElement).getAttribute('href') || null,
				anchor: (el as HTMLElement).dataset.anchor || null,
				tooltip: (el as HTMLElement).getAttribute('title') || null,
			}),
		},
	],
	toDOM: (mark) => [
		'a',
		{
			'data-docx-link': '1',
			href: mark.attrs.href || (mark.attrs.anchor ? `#${safeAttrValue(mark.attrs.anchor)}` : ''),
			...(mark.attrs.anchor ? { 'data-anchor': mark.attrs.anchor } : {}),
			...(mark.attrs.tooltip ? { title: mark.attrs.tooltip } : {}),
			class: 'dve-link',
		},
		0,
	],
};
