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
	// Callers forward optional lookups directly; `undefined` means "no such style/catalog".
	paragraphStyleId?: string | undefined;
	paragraphCatalog?: ParagraphStyleCatalog | undefined;
	runCatalog?: RunStyleCatalog | undefined;
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
	const docDefaults = context.runCatalog?.docDefaults ?? {};
	const paragraphChain = paragraphStyleRunChain(
		context.paragraphStyleId,
		context.paragraphCatalog,
		context.runCatalog,
	);
	const characterChain = characterStyleChain(run.style, context.runCatalog);
	const tableLevels = context.tableStyleRun ? [context.tableStyleRun] : [];
	const levels: RunFormatting[] = [
		docDefaults,
		...paragraphChain,
		...tableLevels,
		...characterChain,
		run,
	];
	const result: RunFormatting = {};
	// Toggle properties (ECMA-376 §17.7.3): each style type (paragraph, table, character) resolves
	// its value through its own basedOn chain; the types then combine by XOR with the document
	// default. Direct formatting on the run sets the value outright, including an explicit off.
	const chainValue = (chain: RunFormatting[], key: (typeof TOGGLE_KEYS)[number]) => {
		for (let index = chain.length - 1; index >= 0; index--)
			if (chain[index][key] !== undefined) return Boolean(chain[index][key]);
		return false;
	};
	for (const key of TOGGLE_KEYS) {
		const direct = run[key];
		const state =
			direct !== undefined
				? Boolean(direct)
				: Boolean(docDefaults[key]) !==
					(chainValue(paragraphChain, key) !==
						(chainValue(tableLevels, key) !== chainValue(characterChain, key)));
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
