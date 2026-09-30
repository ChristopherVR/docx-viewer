// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';
import { children, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { fractionToThemeByte } from './theme-color.js';
import { writeLigatures } from './ligatures.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}
/** On, explicitly off (`w:val="0"`, cancelling a style), or absent. */
function setToggle(
	doc: XmlDocument,
	props: XmlElement,
	local: string,
	value: boolean | undefined,
): void {
	removeChildren(props, local);
	if (value === undefined) return;
	const element = makeW(doc, local);
	if (!value) setAttribute(element, 'val', '0');
	props.appendChild(element);
}

/** Writes the run properties added for character styles, theming and table fidelity. */
export function setExtendedRunProperties(
	doc: XmlDocument,
	props: XmlElement,
	run: TextRun,
	base?: TextRun,
): void {
	const changed = (key: keyof TextRun): boolean => !base || run[key] !== base[key];
	if (changed('ligatures')) writeLigatures(doc, props, run.ligatures);
	if (changed('style')) {
		removeChildren(props, 'rStyle');
		if (run.style) {
			const style = makeW(doc, 'rStyle');
			setAttribute(style, 'val', run.style);
			props.insertBefore(style, props.firstChild);
		}
	}
	if (changed('caps')) setToggle(doc, props, 'caps', run.caps);
	if (changed('smallCaps')) setToggle(doc, props, 'smallCaps', run.smallCaps);
	if (changed('doubleStrike')) setToggle(doc, props, 'dstrike', run.doubleStrike);
	if (changed('vanish')) setToggle(doc, props, 'vanish', run.vanish);
	for (const [key, local] of [
		['textScalePercent', 'w'],
		['kerningHalfPoints', 'kern'],
		['positionHalfPoints', 'position'],
	] as const) {
		if (!changed(key)) continue;
		removeChildren(props, local);
		if (run[key] === undefined) continue;
		const element = makeW(doc, local);
		setAttribute(element, 'val', String(run[key]));
		props.appendChild(element);
	}
	if (changed('characterSpacingTwips')) {
		removeChildren(props, 'spacing');
		if (run.characterSpacingTwips !== undefined) {
			const spacing = makeW(doc, 'spacing');
			setAttribute(spacing, 'val', String(Math.round(run.characterSpacingTwips)));
			props.appendChild(spacing);
		}
	}
	if (changed('shadingFill') || changed('shadingThemeFill')) {
		removeChildren(props, 'shd');
		if (run.shadingFill || run.shadingThemeFill) {
			const shd = makeW(doc, 'shd');
			setAttribute(shd, 'val', 'clear');
			setAttribute(shd, 'fill', run.shadingFill ? run.shadingFill.replace(/^#/, '') : 'auto');
			if (run.shadingThemeFill) {
				setAttribute(shd, 'themeFill', run.shadingThemeFill.token);
				if (run.shadingThemeFill.tint !== undefined)
					setAttribute(shd, 'themeFillTint', fractionToThemeByte(run.shadingThemeFill.tint));
				if (run.shadingThemeFill.shade !== undefined)
					setAttribute(shd, 'themeFillShade', fractionToThemeByte(run.shadingThemeFill.shade));
			}
			props.appendChild(shd);
		}
	}
}
