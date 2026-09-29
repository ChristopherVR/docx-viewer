import type { Block, DocumentModel, Paragraph } from '@christophervr/docx-core';

export interface DocumentStats {
	paragraphs: number;
	tables: number;
	sections: number;
	characters: number;
	charactersWithSpaces: number;
	comments: number;
	revisions: number;
}

/** Text a run shows: field instructions and markers are not text, deleted runs are gone. */
function visibleText(paragraph: Paragraph): string {
	return paragraph.runs
		.filter((run) => !run.fieldCode && !run.fieldChar)
		.filter((run) => run.revision?.kind !== 'delete' && run.revision?.kind !== 'moveFrom')
		.map((run) => run.text)
		.join('');
}

function paragraphsOf(blocks: readonly Block[]): Paragraph[] {
	return blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
}

/** Facts for the File > Info page: counts of what the model holds, not layout results. */
export function documentStats(model: DocumentModel): DocumentStats {
	const paragraphs = paragraphsOf(model.blocks);
	const text = paragraphs.map(visibleText).join('');
	const revisions = paragraphs.reduce(
		(count, paragraph) =>
			count +
			paragraph.runs.filter((run) => run.revision).length +
			(paragraph.markRevision ? 1 : 0),
		0,
	);
	return {
		paragraphs: paragraphs.length,
		tables: model.blocks.filter((block) => block.type === 'table').length,
		sections: model.sections?.length ?? 1,
		characters: [...text.replace(/\s/g, '')].length,
		charactersWithSpaces: [...text].length,
		comments: model.comments?.length ?? 0,
		revisions,
	};
}

/**
 * The document as plain text for File > Export: one line per paragraph, table cells separated by
 * tabs. Formatting, pictures and notes are not exported.
 */
export function plainText(model: DocumentModel): string {
	return model.blocks
		.map((block) =>
			block.type === 'paragraph'
				? visibleText(block)
				: block.rows
						.map((row) => row.map((cell) => cell.paragraphs.map(visibleText).join(' ')).join('\t'))
						.join('\n'),
		)
		.join('\n');
}
