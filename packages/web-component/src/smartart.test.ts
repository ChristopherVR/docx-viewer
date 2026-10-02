// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadDocx, type DocxDiagram, type TextRun } from 'docx-core';
import { EditorState } from 'prosemirror-state';
import { docToModel, modelToDoc } from './model-adapter';
import { appendInlineNode, runToInlineNodes } from './run-adapter';
import { schema } from './schema';
import { diagramLabel, diagramNotes, parseDiagram, smartArtNodeView } from './smartart-node-view';
import { themeDrawing } from './smartart-theme';

// Built by ooxml-core (see tests/support/smartart-fixture.md for provenance).
const fixture = () =>
	new Uint8Array(readFileSync(join(process.cwd(), 'tests/support/smartart.docx')));

async function diagramRuns() {
	const { model } = await loadDocx(fixture());
	const runs = model.blocks.flatMap((block) =>
		block.type === 'paragraph' ? block.runs.filter((run) => run.image?.diagram) : [],
	);
	return { model, runs };
}

describe('SmartArt in the editor model', () => {
	it('keeps run.image.diagram through the node attributes and back', async () => {
		const { runs } = await diagramRuns();
		expect(runs).toHaveLength(2);
		for (const run of runs) {
			const nodes = runToInlineNodes(run);
			expect(nodes).toHaveLength(1);
			expect(nodes[0]!.type).toBe(schema.nodes.image);
			expect(parseDiagram(nodes[0]!.attrs.diagram)).toEqual(run.image!.diagram);
			const back: TextRun[] = [];
			appendInlineNode(back, nodes[0]!);
			expect(back[0]!.image).toEqual(run.image);
			// The writer keeps a placeholder drawing's XML only when the run serializes identically.
			expect(JSON.stringify(back[0]!.image)).toBe(JSON.stringify(run.image));
		}
	});

	it('survives a whole-document round trip, an edit elsewhere, and undo of the node', async () => {
		const { model } = await diagramRuns();
		const doc = modelToDoc(model);
		const state = EditorState.create({ schema, doc });
		const again = docToModel(state.doc, model);
		expect(again.blocks).toEqual(model.blocks);
		// Typing in another paragraph leaves both diagrams byte-for-byte equal in the model.
		let target = 0;
		state.doc.descendants((node, pos) => {
			if (!target && node.isTextblock && node.textContent && !node.firstChild?.attrs.diagram)
				target = pos + 1;
		});
		const edited = state.apply(state.tr.insertText('X', target));
		const result = docToModel(edited.doc, model);
		const diagrams = (blocks: typeof model.blocks) =>
			blocks.flatMap((block) =>
				block.type === 'paragraph' ? block.runs.flatMap((run) => run.image?.diagram ?? []) : [],
			);
		expect(diagrams(result.blocks)).toEqual(diagrams(model.blocks));
		expect(diagrams(result.blocks)).toHaveLength(2);
	});

	it('marks the placeholder DOM so a copy can be told apart', async () => {
		const { runs } = await diagramRuns();
		const node = runToInlineNodes(runs[0]!)[0]!;
		const spec = schema.nodes.image.spec.toDOM!(node) as [string, Record<string, string>];
		expect(spec[1]['data-docx-smartart']).toBe('1');
		expect(node.attrs.unsupported).toBe('SmartArt');
	});
});

describe('SmartArt node view', () => {
	it('mounts office-ui-smartart with the cached drawing, sized to the run, with an accessible name', async () => {
		const { runs, model } = await diagramRuns();
		const node = runToInlineNodes(runs[0]!)[0]!;
		const diagram = parseDiagram(node.attrs.diagram)!;
		const view = smartArtNodeView(node, diagram, model.theme);
		const dom = view.dom as HTMLElement;
		document.body.append(dom);
		expect(dom.style.width).toBe(`${node.attrs.widthPx}px`);
		expect(dom.getAttribute('aria-roledescription')).toBe('SmartArt diagram');
		expect(dom.getAttribute('aria-label')).toBe(diagramLabel(node, diagram));
		expect(dom.tabIndex).toBe(0);
		const element = dom.querySelector('office-ui-smartart')!;
		const svg = element.shadowRoot!.querySelector('svg')!;
		expect(svg.getAttribute('role')).toBe('img');
		expect(svg.children.length).toBe(diagram.drawing!.shapes.length);
		expect(svg.textContent).toMatch(/\S/);
		// The frame, not the shapes' bounding box, sets the scale.
		expect(svg.getAttribute('viewBox')).toBe(
			`0 0 ${diagram.extentEmu.width / 9525} ${diagram.extentEmu.height / 9525}`,
		);
		expect(dom.querySelector('.dve-smartart-notes')!.textContent).toMatch(/not recomputed/);
		expect(view.update?.(node, [], {} as never)).toBe(false);
	});

	it('shows a labelled box with the node text when there is no cached drawing', async () => {
		const { runs } = await diagramRuns();
		const base = runs[0]!.image!.diagram!;
		const { drawing: _drawing, ...rest } = base;
		const diagram: DocxDiagram = { ...rest, rendering: 'placeholder' };
		const node = runToInlineNodes({
			...runs[0]!,
			image: { ...runs[0]!.image!, diagram },
		})[0]!;
		const dom = smartArtNodeView(node, diagram, undefined).dom as HTMLElement;
		expect(dom.querySelector('office-ui-smartart')).toBeNull();
		expect(dom.dataset.rendering).toBe('placeholder');
		const items = [...dom.querySelectorAll('.dve-smartart-placeholder li')].map(
			(li) => li.textContent,
		);
		expect(items.length).toBeGreaterThan(0);
		expect(items).toEqual(diagram.nodes.filter((entry) => entry.text).map((entry) => entry.text));
	});

	it('lists issues and approximations without claiming fidelity', async () => {
		const { runs } = await diagramRuns();
		const diagram: DocxDiagram = {
			...runs[0]!.image!.diagram!,
			issues: [{ code: 'missing-part', message: 'colors part is absent' }],
		};
		const notes = diagramNotes(diagram, { approximatedGeometries: ['star5'], flattened3d: 2 });
		expect(notes.join('\n')).toContain('missing-part: colors part is absent');
		expect(notes.join('\n')).toContain('star5');
		expect(notes.join('\n')).toContain('2 shape(s) with 3D drawn flat');
		expect(notes.at(-1)).toMatch(/not Word-identical/);
	});
});

describe('SmartArt colours', () => {
	it('resolves scheme colours and their transforms through the document theme', async () => {
		const { runs, model } = await diagramRuns();
		const drawing = runs[0]!.image!.diagram!.drawing!;
		const themed = themeDrawing(drawing, model.theme);
		const json = JSON.stringify(themed.drawing);
		expect(json).not.toContain('"kind":"scheme"');
		expect(JSON.stringify(drawing)).toContain('"kind":"scheme"');
		expect(Object.keys(themed.schemeColors)).toContain('accent1');
	});
});
