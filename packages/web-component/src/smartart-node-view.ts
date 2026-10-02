import type { DocxDiagram, ThemeCatalog } from 'docx-core';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { NodeView } from 'prosemirror-view';
import { defineSmartArt } from 'ooxml-ui/smartart';
import { placementClass } from './inline-content-schema';
import { themeDrawing } from './smartart-theme';

const EMU_PER_PX = 9525;

/** What the shared renderer reports (`office-smartart-render`), as far as this view reads it. */
interface RenderReport {
	approximatedGeometries?: string[];
	approximatedFills?: string[];
	flattened3d?: number;
}

type SmartArtElement = HTMLElement & {
	drawing: unknown;
	schemeColors: Record<string, string>;
};

/** Registers `<office-ui-smartart>` once; a no-op without a DOM (server rendering). */
export function registerSmartArt(): void {
	defineSmartArt();
}

export function parseDiagram(value: unknown): DocxDiagram | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		const diagram: unknown = JSON.parse(value);
		return typeof diagram === 'object' && diagram !== null ? (diagram as DocxDiagram) : undefined;
	} catch {
		return undefined;
	}
}

/** The accessible name: alt text, then title, then the object name, then a generic label. */
export function diagramLabel(node: ProseMirrorNode, diagram: DocxDiagram): string {
	return (
		String(node.attrs.altText || '') ||
		String(node.attrs.title || '') ||
		diagram.name ||
		'SmartArt diagram'
	);
}

/** Honest, human-readable notes shown beside the diagram (never a fidelity claim). */
export function diagramNotes(
	diagram: DocxDiagram,
	report: RenderReport | undefined,
	themed?: { unappliedTransforms: string[]; alphaDropped: boolean },
): string[] {
	const notes = [diagram.notice];
	for (const issue of diagram.issues) notes.push(`${issue.code}: ${issue.message}`);
	if (report?.approximatedGeometries?.length)
		notes.push(`Shapes drawn as rectangles: ${report.approximatedGeometries.join(', ')}.`);
	if (report?.approximatedFills?.length)
		notes.push(`Fills approximated: ${report.approximatedFills.join(', ')}.`);
	if (report?.flattened3d) notes.push(`${report.flattened3d} shape(s) with 3D drawn flat.`);
	if (themed?.alphaDropped) notes.push('Transparent colours are drawn opaque.');
	if (themed?.unappliedTransforms.length)
		notes.push(`Colour transforms not applied: ${themed.unappliedTransforms.join(', ')}.`);
	notes.push('Shown read-only from the saved drawing; not recomputed and not Word-identical.');
	return notes.filter(Boolean);
}

function box(node: ProseMirrorNode, diagram: DocxDiagram): HTMLElement {
	const dom = document.createElement('span');
	dom.className = ['dve-smartart', placementClass(node.attrs.placement)].filter(Boolean).join(' ');
	dom.dataset.docxSmartart = '1';
	dom.dataset.placement = diagram.placement;
	dom.dataset.rendering = diagram.rendering;
	dom.tabIndex = 0;
	dom.setAttribute('role', 'group');
	dom.setAttribute('aria-roledescription', 'SmartArt diagram');
	dom.setAttribute('aria-label', diagramLabel(node, diagram));
	dom.style.width = `${node.attrs.widthPx}px`;
	dom.style.height = `${node.attrs.heightPx}px`;
	return dom;
}

function placeholder(diagram: DocxDiagram): HTMLElement {
	const body = document.createElement('span');
	body.className = 'dve-smartart-placeholder';
	const heading = document.createElement('span');
	heading.className = 'dve-smartart-heading';
	heading.textContent = 'SmartArt diagram (no drawing to show)';
	body.append(heading);
	const list = document.createElement('ul');
	for (const entry of diagram.nodes) {
		if (!entry.text) continue;
		const item = document.createElement('li');
		item.textContent = entry.text;
		list.append(item);
	}
	body.append(list);
	return body;
}

let infoCounter = 0;

/**
 * Node view for an image node that carries a SmartArt `diagram`: mounts `<office-ui-smartart>`
 * with the cached drawing (theme colours resolved), scaled from the diagram frame to the run's
 * size. Read-only: the diagram parts are never edited, and the node stays an atom.
 */
export function smartArtNodeView(
	node: ProseMirrorNode,
	diagram: DocxDiagram,
	theme: ThemeCatalog | undefined,
): NodeView {
	registerSmartArt();
	const dom = box(node, diagram);
	const info = document.createElement('button');
	info.type = 'button';
	info.className = 'dve-smartart-info';
	info.textContent = 'i';
	info.setAttribute('aria-label', 'About this SmartArt rendering');
	info.setAttribute('aria-expanded', 'false');
	const panel = document.createElement('span');
	panel.className = 'dve-smartart-notes';
	panel.id = `dve-smartart-notes-${++infoCounter}`;
	panel.setAttribute('role', 'status');
	panel.hidden = true;
	info.setAttribute('aria-controls', panel.id);
	info.addEventListener('click', (event) => {
		event.preventDefault();
		panel.hidden = !panel.hidden;
		info.setAttribute('aria-expanded', String(!panel.hidden));
	});

	let themed: ReturnType<typeof themeDrawing> | undefined;
	const writeNotes = (report?: RenderReport) => {
		const list = document.createElement('ul');
		for (const note of diagramNotes(diagram, report, themed)) {
			const item = document.createElement('li');
			item.textContent = note;
			list.append(item);
		}
		panel.replaceChildren(list);
		dom.title = diagram.notice;
	};

	const drawable =
		diagram.rendering === 'cached-drawing' && (diagram.drawing?.shapes.length ?? 0) > 0;
	if (drawable && diagram.drawing) {
		themed = themeDrawing(diagram.drawing, theme);
		const element = document.createElement('office-ui-smartart') as SmartArtElement;
		element.className = 'dve-smartart-drawing';
		element.setAttribute('label', diagramLabel(node, diagram));
		const frameWidth = diagram.extentEmu.width / EMU_PER_PX;
		const frameHeight = diagram.extentEmu.height / EMU_PER_PX;
		element.addEventListener('office-smartart-render', (event) => {
			// The renderer fits the shapes' bounding box; scale the whole frame to the run size instead.
			const svg = element.shadowRoot?.querySelector('svg');
			if (svg && frameWidth > 0 && frameHeight > 0) {
				svg.setAttribute('viewBox', `0 0 ${frameWidth} ${frameHeight}`);
				svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
				svg.style.height = '100%';
			}
			writeNotes((event as CustomEvent<RenderReport>).detail);
		});
		writeNotes();
		element.schemeColors = themed.schemeColors;
		element.drawing = themed.drawing;
		dom.append(element);
	} else {
		dom.classList.add('dve-smartart-box');
		dom.append(placeholder(diagram));
		writeNotes();
	}
	dom.append(info, panel);
	return {
		dom,
		selectNode: () => dom.classList.add('dve-smartart-selected'),
		deselectNode: () => dom.classList.remove('dve-smartart-selected'),
		stopEvent: (event) => event.target === info,
		ignoreMutation: () => true,
		// Nothing about a diagram is edited in place; a changed node is rebuilt.
		update: () => false,
	};
}
