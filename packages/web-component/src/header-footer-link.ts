import type { DocumentModel, HeaderFooterContent, HeaderFooterSlots } from 'docx-core';
import {
	nextPartName,
	newHeaderFooterId,
	type HeaderFooterKind,
	type HeaderFooterSlot,
} from './header-footer-commands';

/** Missing references inherit independently for each story, as in Word. */
export function effectiveHeaderFooter(
	model: DocumentModel,
	index: number,
	kind: HeaderFooterKind,
	slot: HeaderFooterSlot,
): HeaderFooterContent | undefined {
	for (let i = index; i >= 0; i--) {
		const content = model.sections?.[i]?.[kind]?.[slot];
		if (content) return content;
	}
	return undefined;
}
export function effectiveSlots(
	model: DocumentModel,
	index: number,
	kind: HeaderFooterKind,
): HeaderFooterSlots {
	return Object.fromEntries(
		['default', 'first', 'even'].flatMap((slot) => {
			const content = effectiveHeaderFooter(model, index, kind, slot as HeaderFooterSlot);
			return content ? [[slot, content]] : [];
		}),
	);
}
export function withHeaderFooterLink(
	model: DocumentModel,
	index: number,
	kind: HeaderFooterKind,
	slot: HeaderFooterSlot,
	linked: boolean,
): DocumentModel {
	if (index < 1 || !model.sections?.[index]) return model;
	const own = model.sections[index]![kind]?.[slot];
	if (linked === !own) return model;
	const content = effectiveHeaderFooter(model, index, kind, slot);
	const copy: HeaderFooterContent = {
		partName: nextPartName(model, kind),
		...(content?.partName ? { sourcePartName: content.sourcePartName ?? content.partName } : {}),
		blocks: content
			? structuredClone(content.blocks)
			: [{ type: 'paragraph', id: newHeaderFooterId(), runs: [] }],
	};
	return {
		...model,
		sections: model.sections.map((section, i) => {
			if (i !== index) return section;
			const slots = { ...section[kind] };
			if (linked) delete slots[slot];
			else slots[slot] = copy;
			return { ...section, [kind]: slots };
		}),
	};
}
