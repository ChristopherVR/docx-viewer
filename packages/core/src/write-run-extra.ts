// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { TextRun } from './model.js';
import { children, makeW, type XmlDocument, type XmlElement, WORD_NS } from './xml.js';
import { fractionToThemeByte } from './theme-color.js';

function setAttribute(element: XmlElement, local: string, value: string): void {
	element.setAttributeNS(WORD_NS, `w:${local}`, value);
}
function removeChildren(element: XmlElement, local: string): void {
	for (const child of children(element, local)) element.removeChild(child);
}
function setToggle(doc: XmlDocument, props: XmlElement, local: string, enabled: boolean): void {
	removeChildren(props, local);
	if (enabled) props.appendChild(makeW(doc, local));
}

/** Writes the run properties added for character styles, theming and table fidelity. */
export function setExtendedRunProperties(
	doc: XmlDocument,
	props: XmlElement,
	run: TextRun,
	base?: TextRun,
): void {
	const changed = (key: keyof TextRun): boolean => !base || run[key] !== base[key];
	if (changed('style')) {
		removeChildren(props, 'rStyle');
		if (run.style) {
			const style = makeW(doc, 'rStyle');
			setAttribute(style, 'val', run.style);
			props.insertBefore(style, props.firstChild);
		}
	}
	if (changed('caps')) setToggle(doc, props, 'caps', run.caps === true);
	if (changed('smallCaps')) setToggle(doc, props, 'smallCaps', run.smallCaps === true);
	if (changed('doubleStrike')) setToggle(doc, props, 'dstrike', run.doubleStrike === true);
	if (changed('vanish')) setToggle(doc, props, 'vanish', run.vanish === true);
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
