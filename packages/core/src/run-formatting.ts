// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { ParagraphStyleCatalog, TextRun } from './model.js';
import type { RunFormatting, RunStyleCatalog } from './run-style-model.js';

/**
 * Word "toggle" run properties (ECMA-376 17.7.3): each level in the formatting
 * hierarchy that specifies the property flips its effective state, regardless
 * of the level's own boolean value — so two ancestor levels that both turn
 * bold "on" cancel back out to "off" (Word's documented double-toggle quirk).
 */
const TOGGLE_KEYS = [
	'bold',
	'italic',
	'caps',
	'smallCaps',
	'strike',
	'doubleStrike',
	'vanish',
] as const satisfies readonly (keyof RunFormatting)[];
const OVERRIDE_KEYS = [
	'underline',
	'underlineStyle',
	'underlineColor',
	'highlight',
	'verticalAlign',
	'fontSize',
	'fontFamily',
	'fontTheme',
	'color',
	'colorTheme',
	'characterSpacingTwips',
	'shadingFill',
	'shadingThemeFill',
] as const satisfies readonly (keyof RunFormatting)[];

function paragraphStyleRunChain(
	styleId: string | undefined,
	paragraphCatalog: ParagraphStyleCatalog | undefined,
	runCatalog: RunStyleCatalog | undefined,
): RunFormatting[] {
	if (!styleId || !paragraphCatalog) return [];
	const chain: RunFormatting[] = [];
	const visited = new Set<string>();
	let current: string | undefined = styleId;
	while (current && paragraphCatalog.styles[current] && !visited.has(current)) {
		visited.add(current);
		const runDefinition = runCatalog?.styles[current];
		if (runDefinition) chain.push(runDefinition.formatting);
		current = paragraphCatalog.styles[current].basedOn;
	}
	return chain.reverse();
}

function characterStyleChain(
	styleId: string | undefined,
	runCatalog: RunStyleCatalog | undefined,
): RunFormatting[] {
	if (!styleId || !runCatalog) return [];
	const chain: RunFormatting[] = [];
	const visited = new Set<string>();
	let current: string | undefined = styleId;
	while (current && runCatalog.styles[current] && !visited.has(current)) {
		visited.add(current);
		chain.push(runCatalog.styles[current].formatting);
		current = runCatalog.styles[current].basedOn;
	}
	return chain.reverse();
}

export interface RunFormattingContext {
	paragraphStyleId?: string;
	paragraphCatalog?: ParagraphStyleCatalog;
	runCatalog?: RunStyleCatalog;
	/** Simple table style run formatting for the cell/region containing this run, if any. */
	tableStyleRun?: RunFormatting;
}

/**
 * Resolves default, style-chain and direct run formatting into one effective
 * layer for display. Direct properties on `run` and the source catalogs are
 * never mutated or flattened; this is purely a rendering-time projection.
 */
export function resolveRunFormatting(
	run: TextRun,
	context: RunFormattingContext = {},
): RunFormatting {
	const levels: RunFormatting[] = [
		context.runCatalog?.docDefaults ?? {},
		...paragraphStyleRunChain(
			context.paragraphStyleId,
			context.paragraphCatalog,
			context.runCatalog,
		),
		...(context.tableStyleRun ? [context.tableStyleRun] : []),
		...characterStyleChain(run.style, context.runCatalog),
		run,
	];
	const result: RunFormatting = {};
	for (const key of TOGGLE_KEYS) {
		let state = false;
		for (const level of levels) if (level[key] !== undefined) state = !state;
		if (state) result[key] = true;
	}
	for (const key of OVERRIDE_KEYS) {
		for (const level of levels) {
			const value = level[key];
			if (value !== undefined) (result as Record<string, unknown>)[key] = value;
		}
	}
	return result;
}
