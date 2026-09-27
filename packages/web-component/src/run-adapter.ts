import type { TextRun } from '@christophervr/docx-core';
import type { Node as ProseMirrorNode, Mark } from 'prosemirror-model';
import { schema } from './schema';

function marksForRun(run: TextRun): Mark[] {
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
	return marks;
}

export function runToInlineNodes(run: TextRun): ProseMirrorNode[] {
	if (run.image) {
		const marks = run.link
			? [
					schema.marks.link.create({
						href: run.link.href ?? null,
						anchor: run.link.anchor ?? null,
						tooltip: run.link.tooltip ?? null,
					}),
				]
			: [];
		return [
			schema.nodes.image.create(
				{
					relId: run.image.relId,
					partName: run.image.partName,
					contentType: run.image.contentType,
					widthPx: run.image.widthPx,
					heightPx: run.image.heightPx,
					altText: run.image.altText ?? null,
					title: run.image.title ?? null,
					anchored: Boolean(run.image.anchored),
					unsupported: run.image.unsupported ?? null,
				},
				null,
				marks,
			),
		];
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

function linkFromMarks(child: ProseMirrorNode): TextRun['link'] {
	const link = propertyOfMark(child, 'link');
	if (!link) return undefined;
	const info: NonNullable<TextRun['link']> = {};
	if (link.attrs.href) info.href = link.attrs.href;
	if (link.attrs.anchor) info.anchor = link.attrs.anchor;
	if (link.attrs.tooltip) info.tooltip = link.attrs.tooltip;
	return info.href || info.anchor ? info : undefined;
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
	if (child.type.name === 'image') {
		runs.push({
			text: '',
			image: {
				relId: child.attrs.relId,
				partName: child.attrs.partName,
				contentType: child.attrs.contentType,
				widthPx: child.attrs.widthPx,
				heightPx: child.attrs.heightPx,
				...(child.attrs.altText ? { altText: child.attrs.altText } : {}),
				...(child.attrs.title ? { title: child.attrs.title } : {}),
				...(child.attrs.anchored ? { anchored: true } : {}),
				...(child.attrs.unsupported ? { unsupported: child.attrs.unsupported } : {}),
			},
			...(linkFromMarks(child) ? { link: linkFromMarks(child) } : {}),
		});
		return;
	}
	if (!child.isText && child.type.name !== 'hardBreak') return;
	const run: TextRun = { text: child.isText ? child.text || '' : '\n' };
	const link = linkFromMarks(child);
	if (link) run.link = link;
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
	const sameLink =
		JSON.stringify(previous?.link ?? null) === JSON.stringify(run.link ?? null);
	if (previous && sameLink && fields.every((field) => previous[field] === run[field]))
		previous.text += run.text;
	else runs.push(run);
}
