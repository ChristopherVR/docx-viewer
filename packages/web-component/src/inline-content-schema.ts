import type { NodeSpec, MarkSpec } from 'prosemirror-model';

const safeAttrValue = (value: unknown): string => String(value ?? '').replace(/[;{}"]/g, '');

/**
 * CSS class approximating a floating picture's wrapping: square/tight/through wrapping floats it
 * left or right, top-and-bottom puts it on its own line, and no wrapping (in front of/behind text)
 * keeps it in line. Offsets beyond roughly half a page width are treated as right-aligned.
 */
export function placementClass(value: unknown): string {
	if (typeof value !== 'string') return '';
	let placement: { wrap?: string; align?: string; offsetXPx?: number };
	try {
		placement = JSON.parse(value);
	} catch {
		return '';
	}
	const side =
		placement.align === 'right' || placement.align === 'outside'
			? 'right'
			: placement.align === 'center'
				? 'center'
				: placement.align
					? 'left'
					: (placement.offsetXPx ?? 0) > 312
						? 'right'
						: 'left';
	if (placement.wrap === 'topAndBottom') return `dve-float-block dve-float-block-${side}`;
	if (placement.wrap === 'square' || placement.wrap === 'tight' || placement.wrap === 'through')
		return side === 'center' ? 'dve-float-block dve-float-block-center' : `dve-float-${side}`;
	return '';
}

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
		/** Floating picture wrapping/position (`PicturePlacement`) as JSON; display only. */
		placement: { default: null },
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
						class: [
							'dve-image',
							node.attrs.anchored ? 'dve-image-anchored' : '',
							placementClass(node.attrs.placement),
						]
							.filter(Boolean)
							.join(' '),
						alt: node.attrs.altText || '',
						title: node.attrs.title || '',
						width: String(node.attrs.widthPx),
						height: String(node.attrs.heightPx),
					},
				],
};

/** Only these schemes are ever placed in a rendered `href`; other targets stay in the model only. */
export function isNavigableHref(href: unknown): href is string {
	return typeof href === 'string' && /^(https?:|mailto:)/i.test(href.trim());
}

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
				href:
					(el as HTMLElement).dataset.href ||
					((el as HTMLElement).dataset.anchor ? null : (el as HTMLElement).getAttribute('href')) ||
					null,
				anchor: (el as HTMLElement).dataset.anchor || null,
				tooltip: (el as HTMLElement).getAttribute('title') || null,
			}),
		},
	],
	toDOM: (mark) => [
		'a',
		{
			'data-docx-link': '1',
			href: isNavigableHref(mark.attrs.href)
				? mark.attrs.href
				: mark.attrs.anchor
					? `#${safeAttrValue(mark.attrs.anchor)}`
					: '',
			...(mark.attrs.href ? { 'data-href': mark.attrs.href } : {}),
			...(mark.attrs.anchor ? { 'data-anchor': mark.attrs.anchor } : {}),
			...(mark.attrs.tooltip ? { title: mark.attrs.tooltip } : {}),
			class: 'dve-link',
		},
		0,
	],
};
