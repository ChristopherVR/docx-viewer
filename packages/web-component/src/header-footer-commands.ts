import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	Paragraph,
	SectionProperties,
} from '@christophervr/docx-core';
import { sectionsOf } from './section-commands';

export type HeaderFooterKind = 'headers' | 'footers';
export type PageNumberAlign = 'left' | 'center' | 'right';

let counter = 0;

/** A block id for new header/footer content, unique within this session and the document. */
export const newHeaderFooterId = (): string =>
	`hf-${Date.now().toString(36)}-${(counter++).toString(36)}`;

/** The first free `word/headerN.xml` / `word/footerN.xml` name among the parts the model uses. */
export function nextPartName(model: DocumentModel, kind: HeaderFooterKind): string {
	const stem = kind === 'headers' ? 'header' : 'footer';
	let highest = 0;
	for (const section of model.sections ?? [])
		for (const content of Object.values(section[kind] ?? {})) {
			const match = new RegExp(`^word/${stem}(\\d+)\\.xml$`).exec(content?.partName ?? '');
			if (match?.[1]) highest = Math.max(highest, Number(match[1]));
		}
	return `word/${stem}${highest + 1}.xml`;
}

/** A paragraph holding Word's PAGE field (`fldSimple`), showing 1 until Print Layout recalculates. */
export function pageNumberParagraph(id: string, align: PageNumberAlign): Paragraph {
	return {
		type: 'paragraph',
		id,
		align,
		runs: [{ text: '1', field: { instr: ' PAGE ', simple: true } }],
	};
}

const hasPageField = (paragraph: Paragraph) =>
	paragraph.runs.some(
		(run) => /\bPAGE\b/i.test(run.field?.instr ?? '') || /\bPAGE\b/i.test(run.fieldCode ?? ''),
	);

/** The section that gets the change: the first, which the editing surface shows. */
function firstSection(model: DocumentModel): SectionProperties[] {
	return sectionsOf(model);
}

function withDefaultSlot(
	model: DocumentModel,
	kind: HeaderFooterKind,
	change: (existing: HeaderFooterContent | undefined) => HeaderFooterContent,
): DocumentModel {
	const sections = firstSection(model).map((section, index) => {
		if (index !== 0) return section;
		const next = change(section[kind]?.default);
		return { ...section, [kind]: { ...section[kind], default: next } };
	});
	return { ...model, sections };
}

/**
 * Insert > Page Number: puts a PAGE field at the given position of the first section's header or
 * footer, creating that part when the document has none. An existing header or footer keeps its
 * content; its page-number paragraph is re-aligned, or one is added at the end.
 */
export function withPageNumber(
	model: DocumentModel,
	position: 'top' | 'bottom',
	align: PageNumberAlign,
	newId: () => string,
): DocumentModel {
	const kind: HeaderFooterKind = position === 'top' ? 'headers' : 'footers';
	const base: DocumentModel = model.sections ? model : { ...model, sections: firstSection(model) };
	return withDefaultSlot(base, kind, (existing) => {
		if (!existing)
			return {
				partName: nextPartName(base, kind),
				blocks: [pageNumberParagraph(newId(), align)],
			};
		const blocks: Block[] = existing.blocks.map((block) => structuredClone(block));
		const page = blocks.find(
			(block): block is Paragraph => block.type === 'paragraph' && hasPageField(block),
		);
		if (page) page.align = align;
		else blocks.push(pageNumberParagraph(newId(), align));
		return { ...existing, blocks };
	});
}

/** Insert > Header or Footer > Blank: creates an empty part when there is none; otherwise a no-op. */
export function withBlankHeaderFooter(
	model: DocumentModel,
	kind: HeaderFooterKind,
	newId: () => string,
): DocumentModel {
	const base: DocumentModel = model.sections ? model : { ...model, sections: firstSection(model) };
	if (base.sections?.[0]?.[kind]?.default) return model;
	return withDefaultSlot(base, kind, () => ({
		partName: nextPartName(base, kind),
		blocks: [{ type: 'paragraph', id: newId(), runs: [] }],
	}));
}
