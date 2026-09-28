import type { TextRun } from '@christophervr/docx-core';
import type { Node as ProseMirrorNode, Mark } from 'prosemirror-model';
import { schema } from './schema';
import { explicitOffFields, extraRunProperties } from './run-extra-mark';
import { marksForRun } from './run-marks';

/** A note reference's number in document order and its label in the document's number format. */
export type NoteNumberLookup = (
	kind: 'footnote' | 'endnote',
	id: string,
) => { number: number; label: string };

export function runToInlineNodes(run: TextRun, noteNumber?: NoteNumberLookup): ProseMirrorNode[] {
	if (run.break) return [schema.nodes.pageBreak.create({ kind: run.break })];
	if (run.fieldChar || run.fieldCode !== undefined) {
		const { text: _text, fieldChar, fieldCode, ...format } = run;
		return [
			schema.nodes.fieldMarker.create({
				kind: fieldChar ?? 'code',
				code: fieldCode ?? null,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
		];
	}
	if (run.noteReference) {
		const { kind, id } = run.noteReference;
		const { text: _text, noteReference: _reference, ...format } = run;
		return [
			schema.nodes.noteReference.create({
				kind,
				id,
				number: noteNumber?.(kind, id).number ?? 1,
				label: noteNumber?.(kind, id).label ?? null,
				format: Object.keys(format).length ? JSON.stringify(format) : null,
			}),
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
					placement: run.image.placement ? JSON.stringify(run.image.placement) : null,
					svgPartName: run.image.svgPartName ?? null,
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
	if (revisionMark) {
		const move =
			typeof revisionMark.attrs.move === 'string'
				? (JSON.parse(revisionMark.attrs.move) as NonNullable<TextRun['revision']>['move'])
				: undefined;
		run.revision = {
			// A move without its move data is recorded as a plain insertion or deletion.
			kind: insertion ? (move ? 'moveTo' : 'insert') : move ? 'moveFrom' : 'delete',
			author: String(revisionMark.attrs.author || ''),
			id: String(revisionMark.attrs.id || ''),
			...(revisionMark.attrs.date ? { date: String(revisionMark.attrs.date) } : {}),
			...(move ? { move } : {}),
		};
	}
	const comment = propertyOfMark(child, 'comment');
	if (comment?.attrs.ids?.length) run.commentIds = [...comment.attrs.ids];
	const field = propertyOfMark(child, 'field');
	if (field)
		run.field = {
			instr: String(field.attrs.instr),
			...(field.attrs.simple ? { simple: true } : {}),
		};
	const characterStyle = propertyOfMark(child, 'characterStyle');
	if (characterStyle?.attrs.id) run.style = String(characterStyle.attrs.id);
	const extra = propertyOfMark(child, 'runProperties');
	if (extra?.attrs.props)
		for (const [key, value] of Object.entries(structuredClone(extra.attrs.props) as TextRun)) {
			// An explicit off never overrides a mark the user applied (bold on text that was unbolded).
			const field = key as keyof TextRun;
			if ((explicitOffFields as readonly string[]).includes(key) && run[field] !== undefined)
				continue;
			setRunField(run, field, value);
		}
}

/** Typed per-field assignment, so a value can only be written to a field that accepts it. */
function setRunField<K extends keyof TextRun>(run: TextRun, field: K, value: TextRun[K]): void {
	run[field] = value;
}

export function appendInlineNode(runs: TextRun[], child: ProseMirrorNode): void {
	if (child.type.name === 'pageBreak') {
		runs.push({ text: '', break: child.attrs.kind === 'column' ? 'column' : 'page' });
		return;
	}
	if (child.type.name === 'fieldMarker') {
		const marker: TextRun =
			child.attrs.kind === 'code'
				? { text: '', fieldCode: String(child.attrs.code ?? '') }
				: { text: '', fieldChar: child.attrs.kind };
		if (typeof child.attrs.format === 'string')
			Object.assign(marker, JSON.parse(child.attrs.format) as Partial<TextRun>);
		runs.push(marker);
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
		if (typeof child.attrs.format === 'string')
			Object.assign(reference, JSON.parse(child.attrs.format) as Partial<TextRun>);
		runs.push(reference);
		return;
	}
	if (child.type.name === 'image') {
		const imageLink = linkFromMarks(child);
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
				...(child.attrs.svgPartName ? { svgPartName: child.attrs.svgPartName } : {}),
				...(typeof child.attrs.placement === 'string'
					? { placement: JSON.parse(child.attrs.placement) }
					: {}),
				...(child.attrs.unsupported ? { unsupported: child.attrs.unsupported } : {}),
			},
			...(imageLink && { link: imageLink }),
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
		!previous.fieldChar &&
		previous.fieldCode === undefined &&
		JSON.stringify(previous.field) === JSON.stringify(run.field) &&
		sameRevision &&
		sameComments &&
		sameExtra &&
		sameLink &&
		fields.every((field) => previous[field] === run[field])
	)
		previous.text += run.text;
	else runs.push(run);
}
