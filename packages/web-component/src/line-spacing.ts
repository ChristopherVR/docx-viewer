export type LineSpacingRule = 'auto' | 'exact' | 'atLeast';

export const lineSpacingOptions: Array<[string, string]> = [
	['inherit', 'Line: style default'],
	['auto:240', '1.0 lines'],
	['auto:276', '1.15 lines'],
	['auto:360', '1.5 lines'],
	['auto:480', '2.0 lines'],
	['auto:600', '2.5 lines'],
	['auto:720', '3.0 lines'],
];

export function lineSpacingValue(attrs: Record<string, unknown>): string {
	const rule =
		attrs.lineSpacingRule === 'exact' || attrs.lineSpacingRule === 'atLeast'
			? attrs.lineSpacingRule
			: 'auto';
	if (!Number.isSafeInteger(attrs.lineSpacingTwips))
		return attrs.lineSpacingRule === 'exact' || attrs.lineSpacingRule === 'atLeast'
			? `rule-only:${attrs.lineSpacingRule}`
			: 'inherit';
	return `${rule}:${Number(attrs.lineSpacingTwips)}`;
}

export function lineSpacingLabel(value: string): string {
	if (value === 'inherit') return 'Style default';
	const ruleOnly = /^rule-only:(exact|atLeast)$/.exec(value);
	if (ruleOnly) return `${ruleOnly[1] === 'exact' ? 'Exact' : 'At least'} rule (no amount)`;
	const match = /^(auto|exact|atLeast):(\d+)$/.exec(value);
	if (!match) return 'Mixed paragraphs';
	const [, rule, rawTwips] = match;
	const twips = Number(rawTwips);
	if (rule === 'auto') {
		const ratios: Record<number, string> = {
			240: '1.0',
			276: '1.15',
			360: '1.5',
			480: '2.0',
			600: '2.5',
			720: '3.0',
		};
		return ratios[twips]
			? `${ratios[twips]} lines`
			: `Automatic ${(twips / 240).toLocaleString(undefined, { maximumFractionDigits: 3 })} lines (${twips}/240)`;
	}
	const points = (twips / 20).toLocaleString(undefined, { maximumFractionDigits: 2 });
	return `${rule === 'exact' ? 'Exact' : 'At least'} ${points} pt`;
}

export function parseLineSpacingValue(
	value: string,
): { rule: LineSpacingRule; twips: number } | null {
	const match = /^(auto|exact|atLeast):(\d+)$/.exec(value);
	if (!match) return null;
	const twips = Number(match[2]);
	return Number.isSafeInteger(twips) && twips >= 0
		? { rule: match[1] as LineSpacingRule, twips }
		: null;
}
