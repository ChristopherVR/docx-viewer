// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { equationDom } from './equation-schema';

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const fraction = `<m:oMath xmlns:m="${M}"><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>`;

describe('read-only Word equations', () => {
	it('renders inline fractions as native accessible MathML without editable text', () => {
		const dom = equationDom(fraction, false);
		expect(dom.contentEditable).toBe('false');
		expect(dom.querySelector('mfrac')).not.toBeNull();
		expect(dom.querySelector('math')?.namespaceURI).toBe('http://www.w3.org/1998/Math/MathML');
		expect(dom.querySelector('math')?.getAttribute('display')).toBe('inline');
		expect(dom.querySelector('math')?.getAttribute('aria-label')).toContain('frac');
		expect(dom.textContent).toBe('x2');
	});

	it('renders display equation paragraphs with block MathML', () => {
		const dom = equationDom(`<m:oMathPara xmlns:m="${M}">${fraction}</m:oMathPara>`, true);
		expect(dom.classList.contains('dve-equation-display')).toBe(true);
		expect(dom.querySelector('math')?.getAttribute('display')).toBe('block');
		expect(dom.querySelector('mfrac')).not.toBeNull();
	});

	it.each(['<broken>', `<m:oMath xmlns:m="${M}"><m:unsupported/></m:oMath>`])(
		'reports unavailable rendering accessibly without exposing source markup: %s',
		(omml) => {
			const dom = equationDom(omml, false);
			expect(dom.textContent).toBe('Equation preview unavailable');
			expect(dom.getAttribute('role')).toBe('math');
			expect(dom.getAttribute('aria-label')).toContain('source is preserved');
			expect(dom.querySelector('math')).toBeNull();
		},
	);

	it('keeps equation characters as text rather than HTML', () => {
		const dom = equationDom(
			`<m:oMath xmlns:m="${M}"><m:r><m:t>&lt;img src=x onerror=alert(1)&gt;</m:t></m:r></m:oMath>`,
			false,
		);
		expect(dom.querySelector('img')).toBeNull();
		expect(dom.textContent).toContain('<img');
	});
});
