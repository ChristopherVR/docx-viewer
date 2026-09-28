// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Pre-save model validation: walks a DocumentModel and reports every value the serializer
// could not write as schema-valid WordprocessingML, so `saveDocx` fails loudly instead of
// emitting a file Word rejects. Rules reuse the generated ECMA-376 guards.
import { isStNumberFormat } from './generated/wml-simple-types.js';
import type { Block, DocumentModel } from './model.js';
import { Checker, DocxModelValidationError, type ValidationIssue } from './validate-issues.js';
import { validateParagraph } from './validate-paragraph.js';
import { validateNumberingCatalog, validateSection, validateTable } from './validate-structure.js';

export { DocxModelValidationError, type ValidationIssue } from './validate-issues.js';

function validateBlocks(c: Checker, blocks: Block[] | undefined): void {
	blocks?.forEach((block, index) => {
		const at = c.at(`[${index}]`);
		if (block.type === 'paragraph') return validateParagraph(at, block);
		validateTable(at, block);
		block.rows.forEach((row, r) =>
			row.forEach((cell, k) =>
				cell.paragraphs.forEach((paragraph, p) =>
					validateParagraph(at.at(`.rows[${r}][${k}].paragraphs[${p}]`), paragraph),
				),
			),
		);
	});
}

/** Lists every schema-invalid value in `model`; an empty list means the writer can serialize it. */
export function validateDocumentModel(model: DocumentModel): ValidationIssue[] {
	const issues: ValidationIssue[] = [];
	const path = (suffix: string) => new Checker(issues, suffix);
	const page = path('page');
	for (const key of ['width', 'height'] as const)
		page.check(
			key,
			model.page?.[key],
			(v) => typeof v === 'number' && Number.isFinite(v) && Math.round(v * 15) > 0,
			'must be a finite, positive pixel size (converted to integer twips)',
		);
	for (const key of ['marginTop', 'marginRight', 'marginBottom', 'marginLeft'] as const)
		page.check(
			key,
			model.page?.[key],
			(v) =>
				typeof v === 'number' &&
				Number.isFinite(v) &&
				(key === 'marginTop' || key === 'marginBottom' || v >= 0),
			'must be a finite pixel size (converted to integer twips; left/right non-negative)',
		);
	validateBlocks(path('blocks'), model.blocks);
	model.sections?.forEach((section, index) => {
		const at = path(`sections[${index}]`);
		validateSection(at, section);
		for (const kind of ['headers', 'footers'] as const)
			for (const [slot, content] of Object.entries(section[kind] ?? {}))
				validateBlocks(at.at(`.${kind}.${slot}.blocks`), content?.blocks);
	});
	for (const kind of ['footnotes', 'endnotes'] as const)
		model[kind]?.forEach((note, index) =>
			validateBlocks(path(`${kind}[${index}].blocks`), note.blocks),
		);
	model.comments?.forEach((comment, index) =>
		path(`comments[${index}]`).dateTime('date', comment.date),
	);
	path('model').enum('footnoteNumFmt', model.footnoteNumFmt, isStNumberFormat, 'ST_NumberFormat');
	path('model').enum('endnoteNumFmt', model.endnoteNumFmt, isStNumberFormat, 'ST_NumberFormat');
	if (model.numberingCatalog)
		validateNumberingCatalog(path('numberingCatalog'), model.numberingCatalog);
	return issues;
}

/** Throws `DocxModelValidationError` listing every issue when `validateDocumentModel` finds any. */
export function assertValidDocumentModel(model: DocumentModel): void {
	const issues = validateDocumentModel(model);
	if (issues.length) throw new DocxModelValidationError(issues);
}
