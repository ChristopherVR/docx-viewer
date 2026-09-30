import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { EditorView } from 'prosemirror-view';
import {
	signedTwipsFromPixels,
	twips,
	twipsFromPixels,
	twipsToPixels,
	type DocumentModel,
	type SectionProperties,
	type Twips,
} from '@christophervr/docx-core';
import { expectDefined } from './defined';

const px = twipsToPixels;

/** The document's sections, creating one from the page settings for documents without any. */
export function sectionsOf(model: DocumentModel): SectionProperties[] {
	if (model.sections?.length) return model.sections;
	const { page } = model;
	return [
		{
			endsAtBlockId: model.blocks.at(-1)?.id ?? '',
			type: 'nextPage',
			pageWidthTwips: twipsFromPixels(page.width),
			pageHeightTwips: twipsFromPixels(page.height),
			orientation: page.width > page.height ? 'landscape' : 'portrait',
			marginTopTwips: signedTwipsFromPixels(page.marginTop),
			marginRightTwips: twipsFromPixels(page.marginRight),
			marginBottomTwips: signedTwipsFromPixels(page.marginBottom),
			marginLeftTwips: twipsFromPixels(page.marginLeft),
			headerDistanceTwips: twips(720),
			footerDistanceTwips: twips(720),
			gutterTwips: twips(0),
			columns: { count: 1, spacingTwips: twips(720), equalWidth: true },
		},
	];
}

/** Index of the top-level block holding the selection start. */
function selectedBlockIndex(view: EditorView): number {
	return view.state.selection.$from.index(0);
}

/** The section containing the selection (Word applies page setup to the current section). */
export function currentSectionIndex(view: EditorView, model: DocumentModel): number {
	const sections = sectionsOf(model);
	const block = selectedBlockIndex(view);
	for (const [index, section] of sections.slice(0, -1).entries()) {
		const end = model.blocks.findIndex((item) => item.id === section.endsAtBlockId);
		if (end >= block) return index;
	}
	return sections.length - 1;
}

/** Returns `model` with section `index` changed; the last section also drives `model.page`. */
export function withSection(
	model: DocumentModel,
	index: number,
	change: (section: SectionProperties) => SectionProperties,
): DocumentModel {
	const sections = sectionsOf(model).map((section, position) =>
		position === index ? change(structuredClone(section)) : section,
	);
	const last = sections.at(-1)!;
	return {
		...model,
		sections,
		page: {
			width: px(last.pageWidthTwips),
			height: px(last.pageHeightTwips),
			marginTop: px(last.marginTopTwips),
			marginRight: px(last.marginRightTwips),
			marginBottom: px(last.marginBottomTwips),
			marginLeft: px(last.marginLeftTwips),
		},
	};
}

const NORMAL_MARGIN = twips(1440);
const MARGINS: Record<string, Twips> = {
	normal: NORMAL_MARGIN,
	narrow: twips(720),
	wide: twips(2160),
};

export function setMargins(model: DocumentModel, index: number, preset: string): DocumentModel {
	if (preset === 'moderate')
		return withSection(model, index, (section) => ({
			...section,
			marginTopTwips: NORMAL_MARGIN,
			marginBottomTwips: NORMAL_MARGIN,
			marginLeftTwips: twips(1080),
			marginRightTwips: twips(1080),
		}));
	const value = MARGINS[preset] ?? NORMAL_MARGIN;
	return withSection(model, index, (section) => ({
		...section,
		marginTopTwips: value,
		marginRightTwips: value,
		marginBottomTwips: value,
		marginLeftTwips: value,
	}));
}

export function setOrientation(
	model: DocumentModel,
	index: number,
	orientation: 'portrait' | 'landscape',
): DocumentModel {
	return withSection(model, index, (section) => {
		const { pageWidthTwips: width, pageHeightTwips: height } = section;
		const long = width >= height ? width : height;
		const short = width >= height ? height : width;
		return {
			...section,
			orientation,
			pageWidthTwips: orientation === 'landscape' ? long : short,
			pageHeightTwips: orientation === 'landscape' ? short : long,
		};
	});
}

export function setColumns(model: DocumentModel, index: number, count: number): DocumentModel {
	return withSection(model, index, (section) => ({
		...section,
		columns: { count, spacingTwips: section.columns.spacingTwips ?? twips(720), equalWidth: true },
	}));
}

/** Layout > Page Setup > Vertical alignment for one section. */
export function setVerticalAlign(
	model: DocumentModel,
	index: number,
	verticalAlign: NonNullable<SectionProperties['verticalAlign']>,
): DocumentModel {
	return withSection(model, index, (section) => {
		const { verticalAlign: _previous, ...rest } = section;
		return verticalAlign === 'top' ? rest : { ...rest, verticalAlign };
	});
}

/** Header & Footer > Different First Page for one section. */
export function setTitlePage(
	model: DocumentModel,
	index: number,
	titlePage: boolean,
): DocumentModel {
	return withSection(model, index, (section) => ({ ...section, titlePage }));
}

/** Page number format, and whether numbering continues or restarts at 1 in this section. */
export function setPageNumbering(
	model: DocumentModel,
	index: number,
	change: { format?: NonNullable<SectionProperties['pageNumbering']>['format']; restart?: boolean },
): DocumentModel {
	return withSection(model, index, (section) => {
		const numbering = { ...section.pageNumbering };
		if (change.format !== undefined) numbering.format = change.format;
		if (change.restart === true) numbering.start = 1;
		if (change.restart === false) delete numbering.start;
		if (numbering.format === 'decimal') delete numbering.format;
		const { pageNumbering: _previous, ...rest } = section;
		return Object.keys(numbering).length ? { ...rest, pageNumbering: numbering } : rest;
	});
}

/**
 * Word's Layout > Breaks > Section Break: the current section ends after the selected paragraph
 * and a new section with the same settings begins; `type` is how the new section starts.
 */
export function insertSectionBreak(
	view: EditorView,
	model: DocumentModel,
	type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage',
): DocumentModel {
	const block = model.blocks[selectedBlockIndex(view)];
	if (!block || block.type !== 'paragraph')
		throw new Error('Place the cursor in a paragraph to insert a section break.');
	if (block.id === model.blocks.at(-1)?.id)
		throw new Error('A section break needs text after it; add a paragraph below first.');
	const sections = sectionsOf(model);
	if (sections.some((section) => section.endsAtBlockId === block.id)) return model;
	const index = currentSectionIndex(view, model);
	const current = expectDefined(sections[index], 'current section');
	const next = [...sections];
	next.splice(
		index,
		1,
		{ ...structuredClone(current), endsAtBlockId: block.id },
		{ ...current, type },
	);
	return { ...model, sections: next };
}

const BREAK_LABEL: Record<SectionProperties['type'], string> = {
	nextPage: 'Section Break (Next Page)',
	continuous: 'Section Break (Continuous)',
	evenPage: 'Section Break (Even Page)',
	oddPage: 'Section Break (Odd Page)',
	nextColumn: 'Section Break (Next Column)',
};

/** Marks paragraphs that end a section, like Word's formatting marks for section breaks. */
export function sectionBreaksPlugin() {
	return new Plugin({
		props: {
			decorations(state) {
				const json = state.doc.attrs.sections;
				const sections: SectionProperties[] = typeof json === 'string' ? JSON.parse(json) : [];
				if (sections.length < 2) return null;
				const labels = new Map<string, string>();
				sections.slice(0, -1).forEach((section, index) => {
					const following = sections[index + 1];
					if (following) labels.set(section.endsAtBlockId, BREAK_LABEL[following.type]);
				});
				const decorations: Decoration[] = [];
				state.doc.forEach((node, offset) => {
					const label = labels.get(String(node.attrs.id));
					if (label)
						decorations.push(
							Decoration.node(offset, offset + node.nodeSize, {
								class: 'dve-section-end',
								'data-section-break': label,
							}),
						);
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}

/** Layout > Line Numbers: off, or on with the restart rule; other line-number values are kept. */
export function setLineNumbering(
	model: DocumentModel,
	index: number,
	mode: 'none' | 'continuous' | 'newPage' | 'newSection',
): DocumentModel {
	return withSection(model, index, (section) => {
		const { lineNumbering: _on, lineNumberSettings: previous, ...rest } = section;
		if (mode === 'none') return rest;
		return {
			...rest,
			lineNumbering: true,
			lineNumberSettings: {
				countBy: previous?.countBy ?? 1,
				start: previous?.start ?? 1,
				restart: mode,
				...(previous?.distanceTwips !== undefined ? { distanceTwips: previous.distanceTwips } : {}),
			},
		};
	});
}
