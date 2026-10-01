import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	Paragraph,
	WatermarkSpec,
} from '@christophervr/docx-core';
import { withBlankHeaderFooter } from './header-footer-commands';
import { sectionsOf } from './section-commands';

const SLOTS = ['default', 'first', 'even'] as const;
const isWatermark = (run: Paragraph['runs'][number]) => Boolean(run.image?.watermark);

/** The first watermark in any header, or undefined. */
export function currentWatermark(model: DocumentModel): WatermarkSpec | undefined {
	for (const section of model.sections ?? [])
		for (const content of Object.values(section.headers ?? {}))
			for (const block of content?.blocks ?? [])
				if (block.type === 'paragraph')
					for (const run of block.runs) if (run.image?.watermark) return run.image.watermark;
	return undefined;
}

/** `blocks` with every watermark run removed and, when `spec` is given, one added to the first paragraph. */
function withWatermarkRun(
	blocks: Block[],
	spec: WatermarkSpec | undefined,
	newId: () => string,
): Block[] {
	const next = structuredClone(blocks).map((block) =>
		block.type === 'paragraph'
			? { ...block, runs: block.runs.filter((run) => !isWatermark(run)) }
			: block,
	);
	if (!spec) return next;
	const run = {
		text: '',
		image: {
			relId: '',
			partName: '',
			contentType: 'application/octet-stream',
			widthPx: 0,
			heightPx: 0,
			unsupported: 'Watermark',
			watermark: spec,
		},
	};
	const index = next.findIndex((block) => block.type === 'paragraph');
	if (index < 0) return [{ type: 'paragraph', id: newId(), runs: [run] }, ...next];
	const paragraph = next[index] as Paragraph;
	next[index] = { ...paragraph, runs: [run, ...paragraph.runs] };
	return next;
}

/**
 * Layout > Watermark: sets (or, with undefined, removes) the text watermark in the document's
 * headers. The first section gets a default header if it has none; every other header story the
 * document already has is updated too, and a part shared by several sections changes once for all.
 */
export function withWatermark(
	model: DocumentModel,
	spec: WatermarkSpec | undefined,
	newId: () => string,
): DocumentModel {
	let base = model;
	if (spec && !sectionsOf(base)[0]?.headers?.default)
		base = withBlankHeaderFooter(base, 'headers', newId, 0, 'default');
	const updated = new Map<string, Block[]>();
	const change = (content: HeaderFooterContent): HeaderFooterContent => {
		const key = content.partName ?? '';
		if (!key) return { ...content, blocks: withWatermarkRun(content.blocks, spec, newId) };
		if (!updated.has(key)) updated.set(key, withWatermarkRun(content.blocks, spec, newId));
		return { ...content, blocks: structuredClone(updated.get(key)!) };
	};
	return {
		...base,
		sections: sectionsOf(base).map((section) => {
			if (!section.headers) return section;
			const headers = { ...section.headers };
			for (const slot of SLOTS) {
				const content = headers[slot];
				if (content) headers[slot] = change(content);
			}
			return { ...section, headers };
		}),
	};
}
