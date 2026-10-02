import type { DocumentModel, Paragraph } from 'docx-core';
import { resolveParagraphFormatting } from 'docx-core';
import type { EditorView } from 'prosemirror-view';
import { findLocalizedControl } from './localization';
import { selectionIsListKind } from './list-commands';

type Alignment = 'left' | 'center' | 'right' | 'justify';

const ALIGN_BUTTONS: Array<[string, Alignment]> = [
	['Align left', 'left'],
	['Align center', 'center'],
	['Align right', 'right'],
	['Justify', 'justify'],
];

/** The alignment a paragraph shows, resolved through its style like Word's own toggle buttons. */
function effectiveAlignment(attrs: Record<string, unknown>, model: DocumentModel): Alignment {
	const direct = Object.fromEntries(
		Object.entries(attrs).filter(
			([key, value]) => value != null && (key !== 'style' || value !== ''),
		),
	);
	const catalog = model.paragraphStyles;
	const resolved = catalog
		? resolveParagraphFormatting(
				{ ...direct, type: 'paragraph', runs: [] } as unknown as Paragraph,
				catalog,
			)
		: direct;
	const align = (resolved as { align?: string }).align;
	if (align === 'center' || align === 'right' || align === 'justify') return align;
	if (align === 'left') return 'left';
	return attrs.direction === 'rtl' ? 'right' : 'left';
}

/**
 * Reflects the selection in the toggle-style Home buttons: which alignment applies (only when every
 * selected paragraph agrees), whether the selection is a bulleted or numbered list, and which
 * clipboard commands have something to act on.
 */
export function syncParagraphToggles(
	toolbar: HTMLElement,
	view: EditorView,
	model: DocumentModel,
): void {
	const { selection, doc } = view.state;
	const alignments = new Set<Alignment>();
	if (selection.empty) alignments.add(effectiveAlignment(selection.$from.parent.attrs, model));
	else
		doc.nodesBetween(selection.from, selection.to, (node) => {
			if (node.type.name === 'paragraph') alignments.add(effectiveAlignment(node.attrs, model));
		});
	for (const [label, value] of ALIGN_BUTTONS)
		findLocalizedControl(toolbar, label)?.setAttribute(
			'aria-pressed',
			String(alignments.size === 1 && alignments.has(value)),
		);
	findLocalizedControl(toolbar, 'Bulleted list')?.setAttribute(
		'aria-pressed',
		String(selectionIsListKind(view, 'bullet', model.numberingCatalog)),
	);
	findLocalizedControl(toolbar, 'Numbered list')?.setAttribute(
		'aria-pressed',
		String(selectionIsListKind(view, 'decimal', model.numberingCatalog)),
	);
	for (const label of ['Cut', 'Copy']) {
		const button = findLocalizedControl<HTMLButtonElement>(toolbar, label);
		if (button && selection.empty) button.disabled = true;
	}
	const cut = findLocalizedControl<HTMLButtonElement>(toolbar, 'Cut');
	if (cut && !view.editable) cut.disabled = true;
}
