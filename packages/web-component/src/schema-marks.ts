import type { MarkSpec } from 'prosemirror-model';
import { isWordHighlightToken, type WordHighlightToken } from '@christophervr/docx-core';
import { reviewMarks } from './review-schema';
import { runPropertiesMark } from './run-extra-mark';
import { linkMarkSpec } from './inline-content-schema';

const safeCssValue = (value: unknown): string => String(value ?? '').replace(/[;{}]/g, '');

function parseFontSize(value: string): number | null {
	const match = /^\s*(\d+(?:\.\d+)?)\s*(pt|px)?\s*$/i.exec(value);
	if (!match) return null;
	const size = Number(match[1]) * (match[2]?.toLowerCase() === 'px' ? 0.75 : 1);
	return Number.isFinite(size) && size > 0 ? size : null;
}

/** Word's highlight palette tokens mapped to stable browser colors. */
export const wordHighlightColors = {
	black: '#000000',
	blue: '#0000ff',
	cyan: '#00ffff',
	green: '#00ff00',
	magenta: '#ff00ff',
	red: '#ff0000',
	yellow: '#ffff00',
	white: '#ffffff',
	darkBlue: '#000080',
	darkCyan: '#008080',
	darkGreen: '#008000',
	darkMagenta: '#800080',
	darkRed: '#800000',
	darkYellow: '#808000',
	darkGray: '#808080',
	lightGray: '#c0c0c0',
} satisfies Record<Exclude<WordHighlightToken, 'none'>, string>;

/** Run formatting marks; key order sets ProseMirror mark precedence. */
export const markSpecs: Record<string, MarkSpec> = {
	bold: { parseDOM: [{ tag: 'strong' }, { tag: 'b' }], toDOM: () => ['strong', 0] },
	italic: { parseDOM: [{ tag: 'em' }, { tag: 'i' }], toDOM: () => ['em', 0] },
	underline: { parseDOM: [{ tag: 'u' }], toDOM: () => ['u', 0] },
	strike: {
		parseDOM: [{ tag: 's' }, { tag: 'del' }, { style: 'text-decoration=line-through' }],
		toDOM: () => ['s', 0],
	},
	highlight: {
		attrs: { color: { default: null } },
		parseDOM: [
			{
				tag: 'span[style*="background-color"]',
				getAttrs: (el) => {
					const color = (el as HTMLElement).style.backgroundColor.toLowerCase();
					const entry = Object.entries(wordHighlightColors).find(([, css]) => {
						const hex = css.slice(1);
						const rgb = `rgb(${parseInt(hex.slice(0, 2), 16)}, ${parseInt(hex.slice(2, 4), 16)}, ${parseInt(hex.slice(4, 6), 16)})`;
						return css.toLowerCase() === color || rgb === color;
					});
					return entry ? { color: entry[0] } : false;
				},
			},
		],
		toDOM: (mark) => {
			const value = String(mark.attrs.color || '');
			const token = isWordHighlightToken(value) && value !== 'none' ? value : null;
			return token
				? ['span', { style: `background-color:${wordHighlightColors[token]}` }, 0]
				: ['span', 0];
		},
	},
	verticalAlign: {
		attrs: { value: { default: 'superscript' } },
		excludes: 'verticalAlign',
		parseDOM: [
			{ tag: 'sup', attrs: { value: 'superscript' } },
			{ tag: 'sub', attrs: { value: 'subscript' } },
		],
		toDOM: (mark) => [mark.attrs.value === 'subscript' ? 'sub' : 'sup', 0],
	},
	language: {
		attrs: {
			language: { default: null },
			eastAsiaLanguage: { default: null },
			bidiLanguage: { default: null },
		},
		parseDOM: [
			{
				tag: 'span[lang], span[data-docx-east-asia-language], span[data-docx-bidi-language]',
				getAttrs: (el) => ({
					language: (el as HTMLElement).getAttribute('lang'),
					eastAsiaLanguage: (el as HTMLElement).dataset.docxEastAsiaLanguage || null,
					bidiLanguage: (el as HTMLElement).dataset.docxBidiLanguage || null,
				}),
			},
		],
		toDOM: (mark) => {
			const attrs: Record<string, string> = {};
			if (mark.attrs.language) attrs.lang = mark.attrs.language;
			if (mark.attrs.eastAsiaLanguage)
				attrs['data-docx-east-asia-language'] = mark.attrs.eastAsiaLanguage;
			if (mark.attrs.bidiLanguage) attrs['data-docx-bidi-language'] = mark.attrs.bidiLanguage;
			return ['span', attrs, 0];
		},
	},
	runRtl: {
		attrs: { value: { default: true } },
		excludes: 'runRtl',
		parseDOM: [
			{ tag: 'span[dir="rtl"]', attrs: { value: true } },
			{ tag: 'span[dir="ltr"]', attrs: { value: false } },
		],
		toDOM: (mark) => ['span', { dir: mark.attrs.value ? 'rtl' : 'ltr' }, 0],
	},
	...reviewMarks,
	runProperties: runPropertiesMark,
	link: linkMarkSpec,
	font: {
		attrs: { family: { default: null }, size: { default: null }, color: { default: null } },
		parseDOM: [
			{
				tag: 'span',
				getAttrs: (el) => {
					const style = (el as HTMLElement).style;
					return {
						family: style.fontFamily.trim() || null,
						size: style.fontSize ? parseFontSize(style.fontSize) : null,
						color: style.color || null,
					};
				},
			},
		],
		toDOM: (mark) => [
			'span',
			{
				style: [
					mark.attrs.family && `font-family:${safeCssValue(mark.attrs.family)}`,
					Number.isFinite(mark.attrs.size) &&
						mark.attrs.size > 0 &&
						`font-size:${mark.attrs.size}pt`,
					mark.attrs.color && `color:${safeCssValue(mark.attrs.color)}`,
				]
					.filter(Boolean)
					.join(';'),
			},
			0,
		],
	},
	/** Direct `w:rStyle` character style reference; the toolbar picker edits it, decorations render it. */
	characterStyle: {
		attrs: { id: { default: null } },
		parseDOM: [
			{
				tag: 'span[data-docx-style]',
				getAttrs: (el) => ({ id: (el as HTMLElement).dataset.docxStyle || null }),
			},
		],
		toDOM: (mark) => ['span', { 'data-docx-style': mark.attrs.id }, 0],
	},
};
