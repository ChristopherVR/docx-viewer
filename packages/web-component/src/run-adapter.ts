import type { TextRun } from '@christophervr/docx-core';
import type { Node as ProseMirrorNode, Mark } from 'prosemirror-model';
import { schema } from './schema';

function marksForRun(run: TextRun): Mark[] {
	const marks: Mark[] = [];
	if (run.bold) marks.push(schema.marks.bold.create());
	if (run.italic) marks.push(schema.marks.italic.create());
	if (run.underline) marks.push(schema.marks.underline.create());
	if (run.strike) marks.push(schema.marks.strike.create());
	if (run.highlight) marks.push(schema.marks.highlight.create({ color: run.highlight }));
	if (run.verticalAlign)
		marks.push(schema.marks.verticalAlign.create({ value: run.verticalAlign }));
	if (
		run.language !== undefined ||
		run.eastAsiaLanguage !== undefined ||
		run.bidiLanguage !== undefined
	)
		marks.push(
			schema.marks.language.create({
				language: run.language ?? null,
				eastAsiaLanguage: run.eastAsiaLanguage ?? null,
				bidiLanguage: run.bidiLanguage ?? null,
			}),
		);
	if (run.rtl !== undefined) marks.push(schema.marks.runRtl.create({ value: run.rtl }));
	if (run.fontFamily || run.fontSize || run.color)
		marks.push(
			schema.marks.font.create({
				family: run.fontFamily || null,
				size: run.fontSize || null,
				color: run.color || null,
			}),
		);
	const revision = run.revision;
	if (revision?.kind === 'insert' || revision?.kind === 'moveTo')
		marks.push(
			schema.marks.insertion.create({
				author: revision.author,
				date: revision.date ?? null,
				id: revision.id,
			}),
		);
	else if (revision?.kind === 'delete' || revision?.kind === 'moveFrom')
		marks.push(
			schema.marks.deletion.create({
				author: revision.author,
				date: revision.date ?? null,
				id: revision.id,
			}),
		);
	if (run.commentIds?.length) marks.push(schema.marks.comment.create({ ids: run.commentIds }));
	return marks;
}

export type NoteNumberLookup = (kind: 'footnote' | 'endnote', id: string) => number;

export function runToInlineNodes(run: TextRun, noteNumber?: NoteNumberLookup): ProseMirrorNode[] {
	if (run.break) return [schema.nodes.pageBreak.create({ kind: run.break })];
	if (run.noteReference) {
		const { kind, id } = run.noteReference;
		return [schema.nodes.noteReference.create({ kind, id, number: noteNumber?.(kind, id) ?? 1 })];
	}
	if (!run.text) return [];
	const marks = marksForRun(run);
	return run.text
		.split(/(\n)/)
		.flatMap((piece) =>
			piece === '\n'
				? [schema.nodes.hardBreak.create(null, null, marks)]
				: piece
					? [schema.text(piece, marks)]
					: [],
		);
}

function propertyOfMark(child: ProseMirrorNode, name: string): Mark | undefined {
	return child.marks.find((mark) => mark.type.name === name);
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
	if (child.type.name === 'pageBreak') {
		runs.push({ text: '', break: child.attrs.kind === 'column' ? 'column' : 'page' });
		return;
	}
	if (child.type.name === 'noteReference') {
		runs.push({
			text: '',
			noteReference: {
				kind: child.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
				id: String(child.attrs.id || ''),
			},
		});
		return;
	}
	if (!child.isText && child.type.name !== 'hardBreak') return;
	const run: TextRun = { text: child.isText ? child.text || '' : '\n' };
	if (propertyOfMark(child, 'bold')) run.bold = true;
	if (propertyOfMark(child, 'italic')) run.italic = true;
	if (propertyOfMark(child, 'underline')) run.underline = true;
	if (propertyOfMark(child, 'strike')) run.strike = true;
	const highlight = propertyOfMark(child, 'highlight');
	if (highlight?.attrs.color) run.highlight = highlight.attrs.color;
	const verticalAlign = propertyOfMark(child, 'verticalAlign');
	if (verticalAlign?.attrs.value) run.verticalAlign = verticalAlign.attrs.value;
	const language = propertyOfMark(child, 'language');
	if (language?.attrs.language != null) run.language = language.attrs.language;
	if (language?.attrs.eastAsiaLanguage != null)
		run.eastAsiaLanguage = language.attrs.eastAsiaLanguage;
	if (language?.attrs.bidiLanguage != null) run.bidiLanguage = language.attrs.bidiLanguage;
	const rtl = propertyOfMark(child, 'runRtl');
	if (rtl) run.rtl = rtl.attrs.value;
	const font = propertyOfMark(child, 'font');
	if (font?.attrs.family) run.fontFamily = font.attrs.family;
	if (font?.attrs.size) run.fontSize = font.attrs.size;
	if (font?.attrs.color) run.color = font.attrs.color;
	const insertion = propertyOfMark(child, 'insertion');
	const deletion = propertyOfMark(child, 'deletion');
	const revisionMark = insertion ?? deletion;
	if (revisionMark)
		run.revision = {
			kind: insertion ? 'insert' : 'delete',
			author: String(revisionMark.attrs.author || ''),
			id: String(revisionMark.attrs.id || ''),
			...(revisionMark.attrs.date ? { date: String(revisionMark.attrs.date) } : {}),
		};
	const comment = propertyOfMark(child, 'comment');
	if (comment?.attrs.ids?.length) run.commentIds = [...comment.attrs.ids];
	const previous = runs.at(-1);
	const fields: (keyof TextRun)[] = [
		'bold',
		'italic',
		'underline',
		'strike',
		'highlight',
		'verticalAlign',
		'language',
		'eastAsiaLanguage',
		'bidiLanguage',
		'rtl',
		'fontFamily',
		'fontSize',
		'color',
	];
	const sameRevision = JSON.stringify(previous?.revision) === JSON.stringify(run.revision);
	const sameComments = JSON.stringify(previous?.commentIds) === JSON.stringify(run.commentIds);
	if (
		previous &&
		!previous.break &&
		!previous.noteReference &&
		sameRevision &&
		sameComments &&
		fields.every((field) => previous[field] === run[field])
	)
		previous.text += run.text;
	else runs.push(run);
}
