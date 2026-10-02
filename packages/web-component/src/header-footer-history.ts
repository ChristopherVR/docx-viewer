import type { Block, SectionProperties } from 'docx-core';
import { sectionLayoutJson, sectionsFromLayout } from './section-layout';
import { DocAttrStep, StepMap } from 'prosemirror-transform';
import type { Node } from 'prosemirror-model';

/** Identity mapping at the document boundary lets history group consecutive part-content edits. */
export class SectionPartsStep extends DocAttrStep {
	constructor(value: string) {
		super('sectionParts', value);
	}
	override getMap(): StepMap {
		return new StepMap([0, 0, 0]);
	}
	override invert(doc: Node): SectionPartsStep {
		return new SectionPartsStep(doc.attrs.sectionParts);
	}
}

type SectionParts = Pick<SectionProperties, 'endsAtBlockId' | 'headers' | 'footers'>;
/** Content snapshots live separately from layout so either can enter the same document history. */
export function sectionPartsJson(sections: SectionProperties[] | undefined): string {
	return JSON.stringify(
		(sections ?? []).map(({ endsAtBlockId, headers, footers }) => ({
			endsAtBlockId,
			...(headers ? { headers } : {}),
			...(footers ? { footers } : {}),
		})),
	);
}
export function restoreSectionParts(
	layout: SectionProperties[],
	json: string,
	blocks: Block[],
): SectionProperties[] {
	const parts = JSON.parse(json) as SectionParts[];
	const base = layout.at(-1);
	if (!base) return layout;
	const sources = parts.map((part, i) => {
		const { headers: _headers, footers: _footers, ...geometry } = layout[i] ?? base;
		return { ...geometry, ...part };
	});
	return sectionsFromLayout(sectionLayoutJson(layout), sources, blocks);
}

/** Prevents an in-place editor from being torn down by its own forwarded transaction. */
export const HEADER_FOOTER_INPUT = 'header-footer-input';
