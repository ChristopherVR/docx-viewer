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
	if (run.fontFamily || run.fontSize || run.color)
		marks.push(
			schema.marks.font.create({
				family: run.fontFamily || null,
				size: run.fontSize || null,
				color: run.color || null,
			}),
		);
	return marks;
}

export function runToInlineNodes(run: TextRun): ProseMirrorNode[] {
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
	const font = propertyOfMark(child, 'font');
	if (font?.attrs.family) run.fontFamily = font.attrs.family;
	if (font?.attrs.size) run.fontSize = font.attrs.size;
	if (font?.attrs.color) run.color = font.attrs.color;
	const previous = runs.at(-1);
	const fields: (keyof TextRun)[] = [
		'bold',
		'italic',
		'underline',
		'strike',
		'highlight',
		'verticalAlign',
		'fontFamily',
		'fontSize',
		'color',
	];
	if (previous && fields.every((field) => previous[field] === run[field]))
		previous.text += run.text;
	else runs.push(run);
}
