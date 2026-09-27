import type { TextRun } from '@christophervr/docx-core';
import type { Node as ProseMirrorNode, Mark } from 'prosemirror-model';
import { schema } from './schema';
import { extraRunProperties } from './run-extra-mark';

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
	if (run.style) marks.push(schema.marks.characterStyle.create({ id: run.style }));
	const extra = extraRunProperties(run);
	if (extra) marks.push(schema.marks.runProperties.create({ props: extra }));
	return marks;
}

export type NoteNumberLookup = (kind: 'footnote' | 'endnote', id: string) => number;

export function runToInlineNodes(run: TextRun, noteNumber?: NoteNumberLookup): ProseMirrorNode[] {
	if (run.break) return [schema.nodes.pageBreak.create({ kind: run.break })];
	if (run.noteReference) {
		const { kind, id } = run.noteReference;
		return [
			schema.nodes.noteReference.create(
				{ kind, id, number: noteNumber?.(kind, id) ?? 1 },
				null,
				marksForRun(run),
			),
		];
	}
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

/** Copies run formatting carried by a node's marks onto `run` (text runs and note references). */
function applyMarkFormatting(run: TextRun, child: ProseMirrorNode): void {
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
	const characterStyle = propertyOfMark(child, 'characterStyle');
	if (characterStyle?.attrs.id) run.style = String(characterStyle.attrs.id);
	const extra = propertyOfMark(child, 'runProperties');
	if (extra?.attrs.props) Object.assign(run, structuredClone(extra.attrs.props));
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
	if (child.type.name === 'pageBreak') {
		runs.push({ text: '', break: child.attrs.kind === 'column' ? 'column' : 'page' });
		return;
	}
	if (child.type.name === 'noteReference') {
		const reference: TextRun = {
			text: '',
			noteReference: {
				kind: child.attrs.kind === 'endnote' ? 'endnote' : 'footnote',
				id: String(child.attrs.id || ''),
			},
		};
		// The reference's own formatting (superscript, FootnoteReference style) rides on its marks.
		applyMarkFormatting(reference, child);
		runs.push(reference);
		return;
	}
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
	applyMarkFormatting(run, child);
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
		'style',
	];
	const sameRevision = JSON.stringify(previous?.revision) === JSON.stringify(run.revision);
	const sameComments = JSON.stringify(previous?.commentIds) === JSON.stringify(run.commentIds);
	const sameLink = JSON.stringify(previous?.link ?? null) === JSON.stringify(run.link ?? null);
	const sameExtra =
		JSON.stringify(previous && extraRunProperties(previous)) ===
		JSON.stringify(extraRunProperties(run));
	if (
		previous &&
		!previous.break &&
		!previous.noteReference &&
		sameRevision &&
		sameComments &&
		sameExtra &&
		sameLink &&
		fields.every((field) => previous[field] === run[field])
	)
		previous.text += run.text;
	else runs.push(run);
}
