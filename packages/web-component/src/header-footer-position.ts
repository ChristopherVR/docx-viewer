import { twips, type DocumentModel } from 'docx-core';
import { sectionsOf } from './section-commands';

/** Match Page Setup's measure range while preserving every other property exactly. */
export function withHeaderFooterDistance(
	model: DocumentModel,
	index: number,
	kind: 'header' | 'footer',
	inches: number,
): DocumentModel {
	const sections = sectionsOf(model);
	if (!Number.isFinite(inches) || inches < 0 || inches > 22 || !sections[index]) return model;
	const key = kind === 'header' ? 'headerDistanceTwips' : 'footerDistanceTwips';
	const distance = twips(Math.round(inches * 1440));
	if ((sections[index]![key] ?? 720) === distance) return model;
	return {
		...model,
		sections: sections.map((section, i) =>
			i === index ? { ...section, [key]: distance } : section,
		),
	};
}
