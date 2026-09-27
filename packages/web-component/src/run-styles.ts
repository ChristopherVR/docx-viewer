import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import {
	resolveRunFormatting,
	resolveThemeColorReference,
	type DocumentModel,
	type RunFormatting,
	type TextRun,
} from '@christophervr/docx-core';
import { appendInlineNode } from './run-adapter';

type Theme = NonNullable<DocumentModel['theme']>;

function themeFont(formatting: RunFormatting, theme: Theme | undefined): string | undefined {
	const role = formatting.fontTheme?.ascii ?? formatting.fontTheme?.hAnsi;
	return role && theme ? theme.fonts[role]?.latin : undefined;
}

const quoteFont = (family: string) => `"${family.replace(/["\\]/g, '')}"`;
const safeColor = (value: string | undefined) =>
	value && /^#[0-9a-f]{6}$/i.test(value) ? value : undefined;

/** CSS for inherited run formatting; direct formatting still renders through the run's own marks. */
export function runFormattingCss(
	formatting: RunFormatting,
	theme?: Theme,
	direct: RunFormatting = {},
): string {
	const css: string[] = [];
	// Word toggle properties cancel out when set at an odd number of levels; a direct toggle that the
	// hierarchy cancels still renders its mark, so the decoration (inside the mark) resets it.
	if (formatting.bold) css.push('font-weight:700');
	else if (direct.bold) css.push('font-weight:400');
	if (formatting.italic) css.push('font-style:italic');
	else if (direct.italic) css.push('font-style:normal');
	const lines = [
		formatting.underline ? 'underline' : '',
		formatting.strike || formatting.doubleStrike ? 'line-through' : '',
	].filter(Boolean);
	if (lines.length) css.push(`text-decoration-line:${lines.join(' ')}`);
	if (formatting.caps) css.push('text-transform:uppercase');
	else if (direct.caps) css.push('text-transform:none');
	if (formatting.smallCaps) css.push('font-variant:small-caps');
	else if (direct.smallCaps) css.push('font-variant:normal');
	if (formatting.fontSize) css.push(`font-size:${formatting.fontSize}pt`);
	const family = formatting.fontFamily ?? themeFont(formatting, theme);
	if (family) css.push(`font-family:${quoteFont(family)}`);
	const color =
		safeColor(formatting.color) ??
		(formatting.colorTheme && theme
			? safeColor(resolveThemeColorReference(formatting.colorTheme, theme))
			: undefined);
	if (color) css.push(`color:${color}`);
	return css.join(';');
}

function runAt(node: ProseMirrorNode): TextRun | undefined {
	const runs: TextRun[] = [];
	appendInlineNode(runs, node);
	return runs[0];
}

/**
 * Renders formatting a run inherits from document defaults, its paragraph style, its character
 * style and the theme. It is display-only: decorations never enter the model or collaboration steps.
 */
export function runStylesPlugin(getModel: () => DocumentModel) {
	return new Plugin({
		props: {
			decorations(state) {
				const model = getModel();
				const runCatalog = model.characterStyles;
				if (!runCatalog && !model.theme) return null;
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos, parent) => {
					if (!node.isText) return true;
					const run = runAt(node);
					if (!run) return false;
					const resolved = resolveRunFormatting(run, {
						runCatalog,
						paragraphCatalog: model.paragraphStyles,
						paragraphStyleId: parent?.attrs.style || undefined,
					});
					const style = runFormattingCss(resolved, model.theme, run);
					const hidden = resolved.vanish === true;
					if (style || hidden)
						decorations.push(
							Decoration.inline(pos, pos + node.nodeSize, {
								...(style ? { style } : {}),
								...(hidden ? { class: 'dve-hidden-text' } : {}),
							}),
						);
					return false;
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}
