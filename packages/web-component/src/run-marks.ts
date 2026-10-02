import type { TextRun } from 'docx-core';
import type { Mark } from 'prosemirror-model';
import { schema } from './schema';
import { extraRunProperties } from './run-extra-mark';

/** ProseMirror marks carrying a text run's character formatting, links, fields and revisions. */
export function marksForRun(run: TextRun): Mark[] {
	const marks: Mark[] = [];
	if (run.link)
		marks.push(
			schema.marks.link.create({
				href: run.link.href ?? null,
				anchor: run.link.anchor ?? null,
				tooltip: run.link.tooltip ?? null,
			}),
		);
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
	if (run.field)
		marks.push(
			schema.marks.field.create({ instr: run.field.instr, simple: Boolean(run.field.simple) }),
		);
	const revision = run.revision;
	const move = revision?.move ? JSON.stringify(revision.move) : null;
	if (revision?.kind === 'insert' || revision?.kind === 'moveTo')
		marks.push(
			schema.marks.insertion.create({
				author: revision.author,
				date: revision.date ?? null,
				id: revision.id,
				move: revision.kind === 'moveTo' ? move : null,
			}),
		);
	else if (revision?.kind === 'delete' || revision?.kind === 'moveFrom')
		marks.push(
			schema.marks.deletion.create({
				author: revision.author,
				date: revision.date ?? null,
				id: revision.id,
				move: revision.kind === 'moveFrom' ? move : null,
			}),
		);
	if (run.commentIds?.length) marks.push(schema.marks.comment.create({ ids: run.commentIds }));
	if (run.style) marks.push(schema.marks.characterStyle.create({ id: run.style }));
	const extra = extraRunProperties(run);
	if (extra) marks.push(schema.marks.runProperties.create({ props: extra }));
	return marks;
}
