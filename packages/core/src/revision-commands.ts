// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { Block, DocumentModel, Paragraph } from './model.js';

export interface RevisionEntry {
	id: string;
	kind: 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'formatChange' | 'paragraphChange';
	author: string;
	date?: string;
	paragraphId: string;
	/** Present for run-level revisions; absent for paragraph mark/format revisions. */
	runIndex?: number;
}

function paragraphsOf(blocks: Block[]): Paragraph[] {
	const paragraphs: Paragraph[] = [];
	for (const block of blocks)
		if (block.type === 'paragraph') paragraphs.push(block);
		else for (const row of block.rows) for (const cell of row) paragraphs.push(...cell.paragraphs);
	return paragraphs;
}

/** Lists every run/paragraph revision in document order, for navigation and enumeration. */
export function listRevisions(model: DocumentModel): RevisionEntry[] {
	const entries: RevisionEntry[] = [];
	for (const paragraph of paragraphsOf(model.blocks)) {
		if (paragraph.markRevision)
			entries.push({ ...paragraph.markRevision, paragraphId: paragraph.id });
		if (paragraph.formatRevision)
			entries.push({ ...paragraph.formatRevision, paragraphId: paragraph.id });
		paragraph.runs.forEach((run, runIndex) => {
			if (run.revision) entries.push({ ...run.revision, paragraphId: paragraph.id, runIndex });
		});
	}
	return entries;
}

export function findRevision(model: DocumentModel, id: string): RevisionEntry | undefined {
	return listRevisions(model).find((entry) => entry.id === id);
}

function mergeIntoNext(blocks: Block[], paragraphId: string): Block[] {
	const index = blocks.findIndex((b) => b.type === 'paragraph' && b.id === paragraphId);
	const next = blocks[index + 1];
	if (index < 0 || !next || next.type !== 'paragraph')
		throw new Error(
			`Cannot resolve the paragraph mark revision on paragraph ${paragraphId}: there is no following paragraph to merge with (merging across a table boundary is not supported).`,
		);
	const current = blocks[index] as Paragraph;
	const merged: Paragraph = { ...next, runs: [...current.runs, ...next.runs] };
	return [...blocks.slice(0, index), merged, ...blocks.slice(index + 2)];
}

/** All runs carrying one revision share the same wrapper `id`; ordered indices for splicing. */
function matchingRunIndices(paragraph: Paragraph, id: string): number[] {
	return paragraph.runs.reduce<number[]>((indices, run, index) => {
		if (run.revision?.id === id) indices.push(index);
		return indices;
	}, []);
}
function updateRuns(
	model: DocumentModel,
	id: string,
	apply: (paragraph: Paragraph, runIndex: number) => void,
): DocumentModel {
	const next = structuredClone(model);
	for (const paragraph of paragraphsOf(next.blocks)) {
		const indices = matchingRunIndices(paragraph, id);
		if (!indices.length) continue;
		for (const index of [...indices].sort((a, b) => b - a)) apply(paragraph, index);
		if (!paragraph.runs.length) paragraph.runs.push({ text: '' });
		return next;
	}
	throw new Error(`No run revision with id ${id} was found.`);
}

function resolveParagraphMark(model: DocumentModel, id: string, keepBreak: boolean): DocumentModel {
	const next = structuredClone(model);
	const paragraph = paragraphsOf(next.blocks).find((p) => p.markRevision?.id === id);
	if (!paragraph) throw new Error(`No paragraph mark revision with id ${id} was found.`);
	if (keepBreak) {
		paragraph.markRevision = undefined;
		return next;
	}
	next.blocks = mergeIntoNext(next.blocks, paragraph.id);
	return next;
}

/** Accepts one revision: keeps insertions, drops deletions, clears format markers. */
export function acceptRevision(model: DocumentModel, id: string): DocumentModel {
	const entry = findRevision(model, id);
	if (!entry) throw new Error(`No revision with id ${id} was found.`);
	if (entry.kind === 'formatChange')
		return updateRuns(model, id, (paragraph, runIndex) => {
			paragraph.runs[runIndex].revision = undefined;
		});
	if (entry.kind === 'paragraphChange') {
		const next = structuredClone(model);
		const paragraph = paragraphsOf(next.blocks).find((p) => p.id === entry.paragraphId)!;
		paragraph.formatRevision = undefined;
		return next;
	}
	if (entry.runIndex === undefined) return resolveParagraphMark(model, id, entry.kind === 'insert');
	if (entry.kind === 'insert' || entry.kind === 'moveTo')
		return updateRuns(model, id, (paragraph, runIndex) => {
			paragraph.runs[runIndex].revision = undefined;
		});
	return updateRuns(model, id, (paragraph, runIndex) => {
		paragraph.runs.splice(runIndex, 1);
	});
}

/** Rejects one revision: drops insertions, restores deletions. Formatting-only changes cannot be reverted. */
export function rejectRevision(model: DocumentModel, id: string): DocumentModel {
	const entry = findRevision(model, id);
	if (!entry) throw new Error(`No revision with id ${id} was found.`);
	if (entry.kind === 'formatChange' || entry.kind === 'paragraphChange')
		throw new Error(
			'Rejecting a formatting-only revision is not supported because its prior formatting snapshot is not modeled; use accept to clear the marker instead.',
		);
	if (entry.runIndex === undefined) return resolveParagraphMark(model, id, entry.kind === 'delete');
	if (entry.kind === 'insert' || entry.kind === 'moveTo')
		return updateRuns(model, id, (paragraph, runIndex) => {
			paragraph.runs.splice(runIndex, 1);
		});
	return updateRuns(model, id, (paragraph, runIndex) => {
		paragraph.runs[runIndex].revision = undefined;
	});
}

function applyAll(
	model: DocumentModel,
	apply: (model: DocumentModel, id: string) => DocumentModel,
	fallbackWarning?: string,
): DocumentModel {
	let current = model;
	let warned = false;
	for (const entry of listRevisions(model).reverse()) {
		if (!findRevision(current, entry.id)) continue;
		try {
			current = apply(current, entry.id);
		} catch (error) {
			if (!fallbackWarning) throw error;
			const next = structuredClone(current);
			const paragraph = paragraphsOf(next.blocks).find((p) => p.id === entry.paragraphId)!;
			if (entry.runIndex !== undefined) paragraph.runs[entry.runIndex].revision = undefined;
			else paragraph.formatRevision = undefined;
			current = next;
			warned = true;
		}
	}
	if (warned && !current.warnings.includes(fallbackWarning!))
		current.warnings.push(fallbackWarning!);
	return current;
}

export const acceptAllRevisions = (model: DocumentModel): DocumentModel =>
	applyAll(model, acceptRevision);
export const rejectAllRevisions = (model: DocumentModel): DocumentModel =>
	applyAll(
		model,
		rejectRevision,
		'Rejecting all changes cleared formatting-change markers without restoring their prior formatting, because the prior formatting snapshot is not modeled.',
	);
