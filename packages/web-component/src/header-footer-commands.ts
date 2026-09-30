import type {
	Block,
	DocumentModel,
	HeaderFooterContent,
	Paragraph,
} from '@christophervr/docx-core';
import { sectionsOf } from './section-commands';

export type HeaderFooterKind = 'headers' | 'footers';
export type PageNumberAlign = 'left' | 'center' | 'right';
export type PageNumberStyle = 'plain' | 'pageOfTotal';
export type HeaderFooterSlot = 'default' | 'first' | 'even';

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
export function pageNumberParagraph(
	id: string,
	align: PageNumberAlign,
	style: PageNumberStyle = 'plain',
): Paragraph {
	return {
		type: 'paragraph',
		id,
		align,
		runs:
			style === 'plain'
				? [{ text: '1', field: { instr: ' PAGE ', simple: true } }]
				: [
						{ text: 'Page ' },
						{ text: '1', field: { instr: ' PAGE ', simple: true } },
						{ text: ' of ' },
						{ text: '1', field: { instr: ' NUMPAGES ', simple: true } },
					],
	};
}

const hasPageField = (paragraph: Paragraph) =>
	paragraph.runs.some(
		(run) => /\bPAGE\b/i.test(run.field?.instr ?? '') || /\bPAGE\b/i.test(run.fieldCode ?? ''),
	);

function enableSlot(
	model: DocumentModel,
	sectionIndex: number,
	slot: HeaderFooterSlot,
): DocumentModel {
	const sections = sectionsOf(model);
	if (slot === 'first' && !sections[sectionIndex]?.titlePage)
		return {
			...model,
			sections: sections.map((section, index) =>
				index === sectionIndex ? { ...section, titlePage: true } : section,
			),
		};
	if (slot === 'even' && !model.evenAndOddHeaders)
		return { ...model, sections, evenAndOddHeaders: true };
	return model.sections ? model : { ...model, sections };
}

function withSlot(
	model: DocumentModel,
	kind: HeaderFooterKind,
	change: (existing: HeaderFooterContent | undefined) => HeaderFooterContent,
	sectionIndex: number,
	slot: HeaderFooterSlot,
	inherit = false,
): DocumentModel {
	let owner = sectionIndex;
	const source = sectionsOf(model);
	while (inherit && owner > 0 && !source[owner]?.[kind]?.[slot]) owner--;
	const existing = source[owner]?.[kind]?.[slot];
	if (!existing) owner = sectionIndex;
	const next = change(existing);
	const sections = sectionsOf(model).map((section, index) => {
		if (index !== owner) return section;
		return { ...section, [kind]: { ...section[kind], [slot]: next } };
	});
	if (!existing?.partName || next.partName !== existing.partName) return { ...model, sections };
	return {
		...model,
		sections: sections.map((section) =>
			!section[kind]
				? section
				: {
						...section,
						[kind]: Object.fromEntries(
							Object.entries(section[kind] ?? {}).map(([slot, content]) => [
								slot,
								content?.partName === existing.partName
									? { ...content, blocks: structuredClone(next.blocks) }
									: content,
							]),
						),
					},
		),
	};
}

/**
 * Insert > Page Number: puts a PAGE field at the given position of the selected section's header or
 * footer, creating that part when the document has none. An existing header or footer keeps its
 * content; its page-number paragraph is re-aligned, or one is added at the end.
 */
export function withPageNumber(
	model: DocumentModel,
	position: 'top' | 'bottom',
	align: PageNumberAlign,
	newId: () => string,
	sectionIndex = 0,
	style?: PageNumberStyle,
	slot: HeaderFooterSlot = 'default',
): DocumentModel {
	const kind: HeaderFooterKind = position === 'top' ? 'headers' : 'footers';
	const base = enableSlot(model, sectionIndex, slot);
	return withSlot(
		base,
		kind,
		(existing) => {
			if (!existing)
				return {
					partName: nextPartName(base, kind),
					blocks: [pageNumberParagraph(newId(), align, style)],
				};
			const blocks: Block[] = existing.blocks.map((block) => structuredClone(block));
			const page = blocks.find(
				(block): block is Paragraph => block.type === 'paragraph' && hasPageField(block),
			);
			if (page) {
				page.align = align;
				// Replace only the generated simple number pattern; keep imported custom text/codes.
				const simpleNumber = page.runs.every(
					(run) =>
						(run.field?.simple && /^(PAGE|NUMPAGES)$/i.test(run.field.instr.trim())) ||
						(!run.field &&
							!run.fieldCode &&
							!run.fieldChar &&
							['Page ', ' of '].includes(run.text)),
				);
				if (style && simpleNumber) page.runs = pageNumberParagraph(page.id, align, style).runs;
				else if (
					style === 'pageOfTotal' &&
					!page.runs.some((run) => /\bNUMPAGES\b/i.test(run.field?.instr ?? run.fieldCode ?? ''))
				) {
					page.runs.push(
						{ text: ' of ' },
						{ text: '1', field: { instr: ' NUMPAGES ', simple: true } },
					);
				}
			} else blocks.push(pageNumberParagraph(newId(), align, style));
			return { ...existing, blocks };
		},
		sectionIndex,
		slot,
		true,
	);
}

/** Insert > Header or Footer > Blank: creates an empty part when there is none; otherwise a no-op. */
export function withBlankHeaderFooter(
	model: DocumentModel,
	kind: HeaderFooterKind,
	newId: () => string,
	sectionIndex = 0,
	slot: HeaderFooterSlot = 'default',
): DocumentModel {
	const base = enableSlot(model, sectionIndex, slot);
	if (base.sections?.[sectionIndex]?.[kind]?.[slot]) return base;
	return withSlot(
		base,
		kind,
		() => ({
			partName: nextPartName(base, kind),
			blocks: [{ type: 'paragraph', id: newId(), runs: [] }],
		}),
		sectionIndex,
		slot,
	);
}
