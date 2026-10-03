import type { NodeSpec } from 'prosemirror-model';
import { convertOmmlToLatex, convertOmmlToMathMl, parseOmml } from 'ooxml-core/math';

const MATH_NS = 'http://www.w3.org/1998/Math/MathML';

/** Read-only native MathML display. The source OMML stays in the atom for saving. */
export function equationDom(omml: string, display: boolean): HTMLElement {
	const span = document.createElement('span');
	span.dataset.docxEquation = '1';
	span.className = display ? 'dve-equation dve-equation-display' : 'dve-equation';
	span.contentEditable = 'false';
	span.title = 'Equation preview (display only; Word layout may differ)';
	try {
		const parsed = parseOmml(omml);
		const latex = convertOmmlToLatex(parsed);
		const markup = convertOmmlToMathMl(parsed);
		const math = new DOMParser().parseFromString(markup, 'application/xml').documentElement;
		if (math.localName !== 'math' || math.namespaceURI !== MATH_NS || !math.textContent?.trim())
			throw new Error('No supported equation content');
		math.setAttribute('display', display ? 'block' : 'inline');
		math.setAttribute('aria-label', latex || 'Equation');
		span.append(document.importNode(math, true));
	} catch {
		span.classList.add('dve-equation-unavailable');
		span.setAttribute('role', 'math');
		span.setAttribute('aria-label', 'Equation preview unavailable; source is preserved');
		span.textContent = 'Equation preview unavailable';
	}
	return span;
}

export const equationNodeSpec: NodeSpec = {
	group: 'inline',
	inline: true,
	atom: true,
	selectable: true,
	attrs: { omml: { default: '' }, display: { default: false }, format: { default: null } },
	toDOM: (node) => equationDom(String(node.attrs.omml), Boolean(node.attrs.display)),
};
