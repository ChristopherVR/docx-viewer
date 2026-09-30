import {
	resolveRunFormatting,
	isLigatures,
	type Ligatures,
	type RunFormatting,
} from '@christophervr/docx-core';
import type { Mark, Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { effectiveFont } from './font-sync';
import { applyVerticalAlign } from './inline-commands';
import { applyFont } from './ribbon-commands';
import { runOf, styleModelOf } from './run-styles';
import { schema } from './schema';
import { toggleFormat, type ToggleKey } from './toggle-commands';
import { batchFontFormat } from './font-format-batch';

export type UnderlineKind = 'none' | 'single' | 'double' | 'dotted' | 'dash' | 'wave';
export type Script = 'none' | 'superscript' | 'subscript';

/** Formatting edited by the Font dialog, in UI units (points instead of twips). */
export interface FontFormat {
	family: string;
	size: number;
	color: string;
	bold: boolean;
	italic: boolean;
	underline: UnderlineKind;
	/** `null` is Word's "Automatic" (the text colour). */
	underlineColor: string | null;
	strike: boolean;
	doubleStrike: boolean;
	script: Script;
	smallCaps: boolean;
	caps: boolean;
	hidden: boolean;
	/** Expanded (positive) or condensed (negative) character spacing in points. */
	spacing: number;
	scale: number;
	/** Baseline displacement in points; positive raises, negative lowers. */
	position: number;
	/** Minimum point size for kerning; zero means off. */
	kerning: number;
	ligatures: Ligatures;
}

/** A selection's format, with `null` for a field whose value differs across the selection. */
export type FontFormatState = { [K in keyof FontFormat]: FontFormat[K] | null };

const FIELDS = [
	'family',
	'size',
	'color',
	'bold',
	'italic',
	'underline',
	'underlineColor',
	'strike',
	'doubleStrike',
	'script',
	'smallCaps',
	'caps',
	'hidden',
	'spacing',
	'scale',
	'position',
	'kerning',
	'ligatures',
] as const satisfies readonly (keyof FontFormat)[];

function underlineKind(resolved: RunFormatting): UnderlineKind {
	if (!resolved.underline) return 'none';
	const style = resolved.underlineStyle ?? '';
	if (/double/i.test(style)) return 'double';
	if (/dot/i.test(style)) return 'dotted';
	if (/dash/i.test(style)) return 'dash';
	if (/wav/i.test(style)) return 'wave';
	return 'single';
}

function formatOf(
	state: EditorState,
	text: ProseMirrorNode,
	paragraph: ProseMirrorNode,
): FontFormat {
	const model = styleModelOf(state);
	const run = runOf(text);
	const resolved: RunFormatting =
		run && model
			? resolveRunFormatting(run, {
					runCatalog: model.characterStyles,
					paragraphCatalog: model.paragraphStyles,
					paragraphStyleId: paragraph.attrs.style || undefined,
				})
			: (run ?? {});
	const { family, size } = effectiveFont(state, text, paragraph);
	return {
		family,
		size,
		color: resolved.color && /^#[0-9a-f]{6}$/i.test(resolved.color) ? resolved.color : '#000000',
		bold: Boolean(resolved.bold),
		italic: Boolean(resolved.italic),
		underline: underlineKind(resolved),
		underlineColor:
			resolved.underlineColor && /^#[0-9a-f]{6}$/i.test(resolved.underlineColor)
				? resolved.underlineColor
				: null,
		strike: Boolean(resolved.strike),
		doubleStrike: Boolean(resolved.doubleStrike),
		script:
			resolved.verticalAlign === 'superscript' || resolved.verticalAlign === 'subscript'
				? resolved.verticalAlign
				: 'none',
		smallCaps: Boolean(resolved.smallCaps),
		caps: Boolean(resolved.caps),
		hidden: Boolean(resolved.vanish),
		spacing: (resolved.characterSpacingTwips ?? 0) / 20,
		scale: resolved.textScalePercent ?? 100,
		position: (resolved.positionHalfPoints ?? 0) / 2,
		kerning: (resolved.kerningHalfPoints ?? 0) / 2,
		ligatures: resolved.ligatures ?? 'none',
	};
}

/** The format of each text piece the selection covers (the caret's own when collapsed). */
function formatsOf(state: EditorState): FontFormat[] {
	const { selection, doc } = state;
	const formats: FontFormat[] = [];
	if (selection.empty) {
		const marks = state.storedMarks ?? selection.$from.marks();
		formats.push(formatOf(state, schema.text('x', marks), selection.$from.parent));
	} else {
		doc.nodesBetween(selection.from, selection.to, (node, _pos, parent) => {
			if (node.isText && parent) formats.push(formatOf(state, node, parent));
		});
		if (!formats.length)
			formats.push(
				formatOf(state, schema.text('x', selection.$from.marks()), selection.$from.parent),
			);
	}
	return formats;
}

/** The selection's font format (or the caret's), with `null` where the selection mixes values. */
export function readFontFormat(state: EditorState): FontFormatState {
	const formats = formatsOf(state);
	const result = {} as Record<string, unknown>;
	for (const field of FIELDS) {
		const first = formats[0]?.[field];
		result[field] = formats.every((format) => format[field] === first) ? first : null;
	}
	return result as FontFormatState;
}

/** Whether every piece shows `key` (true), none does (false), or the selection is mixed (null). */
function shows(state: EditorState, key: ToggleKey): boolean | null {
	const values = formatsOf(state).map((format) =>
		key === 'underline' ? format.underline !== 'none' : format[key],
	);
	if (values.every(Boolean)) return true;
	return values.some(Boolean) ? null : false;
}

/** Drives Word's toggle command until the selection shows `desired`, handling mixed selections. */
function setToggle(view: EditorView, key: ToggleKey, desired: boolean): void {
	for (let attempt = 0; attempt < 2; attempt++) {
		if (shows(view.state, key) === desired) return;
		toggleFormat(key)(view.state, view.dispatch, view);
	}
}

type ExtraPatch = Record<string, unknown>;

/** Merges `patch` into each text run's opaque run-properties mark; `undefined` removes a property. */
function patchExtraProps(view: EditorView, patch: ExtraPatch): void {
	const type = schema.marks.runProperties;
	const next = (marks: readonly Mark[]): Mark[] => {
		const existing = type.isInSet(marks);
		const props: ExtraPatch = { ...((existing?.attrs.props as ExtraPatch | null) ?? {}) };
		for (const [key, value] of Object.entries(patch))
			if (value === undefined) delete props[key];
			else props[key] = value;
		const others = marks.filter((mark) => mark.type !== type);
		return Object.keys(props).length ? [...others, type.create({ props })] : others;
	};
	const { state } = view;
	if (state.selection.empty) {
		view.dispatch(
			state.tr.setStoredMarks(next(state.storedMarks ?? state.selection.$from.marks())),
		);
		return;
	}
	const tr = state.tr;
	state.doc.nodesBetween(state.selection.from, state.selection.to, (node, pos) => {
		if (!node.isText) return;
		const from = Math.max(pos, state.selection.from);
		const to = Math.min(pos + node.nodeSize, state.selection.to);
		tr.removeMark(from, to, type);
		const wanted = type.isInSet(next(node.marks));
		if (wanted) tr.addMark(from, to, wanted);
	});
	view.dispatch(tr);
}

/** Applies the fields present in `changes` to the selection; untouched fields keep their values. */
export function applyFontFormat(view: EditorView, changes: Partial<FontFormat>): void {
	if (!view.editable) return;
	batchFontFormat(view, (draft) => applyFontFields(draft, changes));
}

function applyFontFields(view: EditorView, changes: Partial<FontFormat>): void {
	if (changes.family !== undefined) applyFont(view, 'family', changes.family);
	if (changes.size !== undefined) applyFont(view, 'size', String(changes.size));
	if (changes.color !== undefined) applyFont(view, 'color', changes.color);
	if (changes.bold !== undefined) setToggle(view, 'bold', changes.bold);
	if (changes.italic !== undefined) setToggle(view, 'italic', changes.italic);
	if (changes.strike !== undefined) setToggle(view, 'strike', changes.strike);
	if (changes.underline !== undefined) {
		setToggle(view, 'underline', changes.underline !== 'none');
		patchExtraProps(view, {
			underlineStyle:
				changes.underline === 'none' || changes.underline === 'single'
					? undefined
					: changes.underline,
		});
	}
	if (changes.script !== undefined) {
		const current = readFontFormat(view.state).script;
		if (current !== changes.script) {
			applyVerticalAlign(view, changes.script === 'none' ? 'baseline' : changes.script);
		}
	}
	const extras: ExtraPatch = {};
	if (isLigatures(changes.ligatures)) extras.ligatures = changes.ligatures;
	if (changes.underlineColor !== undefined)
		extras.underlineColor = changes.underlineColor ?? undefined;
	if (changes.doubleStrike !== undefined) extras.doubleStrike = changes.doubleStrike;
	if (changes.smallCaps !== undefined) extras.smallCaps = changes.smallCaps;
	if (changes.caps !== undefined) extras.caps = changes.caps;
	if (changes.hidden !== undefined) extras.vanish = changes.hidden;
	if (changes.spacing !== undefined)
		extras.characterSpacingTwips = Math.round(changes.spacing * 20);
	if (
		changes.scale !== undefined &&
		Number.isInteger(changes.scale) &&
		changes.scale >= 0 &&
		changes.scale <= 600
	)
		extras.textScalePercent = changes.scale;
	if (changes.position !== undefined && Number.isFinite(changes.position))
		extras.positionHalfPoints =
			Math.sign(changes.position) * Math.round(Math.abs(changes.position) * 2);
	if (changes.kerning !== undefined && Number.isFinite(changes.kerning) && changes.kerning >= 0)
		extras.kerningHalfPoints = Math.round(changes.kerning * 2);
	if (Object.keys(extras).length) patchExtraProps(view, extras);
}
