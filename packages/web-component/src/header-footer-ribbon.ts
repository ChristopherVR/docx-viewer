import type { DocumentModel } from 'docx-core';
import { group, stack, tool, spinner } from './ribbon-parts';
import type { HeaderFooterKind, HeaderFooterSlot } from './header-footer-commands';

export interface HeaderFooterContext {
	kind: HeaderFooterKind;
	slot: HeaderFooterSlot;
	index: number;
}
export function buildHeaderFooterPanel(panel: HTMLElement): void {
	panel.append(
		group(
			'Navigation',
			stack(
				tool(
					'Go to Header',
					'header',
					{ type: 'headerFooterNavigate', target: 'header' },
					{ inline: true },
				),
				tool(
					'Go to Footer',
					'footer',
					{ type: 'headerFooterNavigate', target: 'footer' },
					{ inline: true },
				),
			),
			stack(
				tool(
					'Previous Section',
					'previous',
					{ type: 'headerFooterNavigate', target: 'previous' },
					{ inline: true },
				),
				tool(
					'Next Section',
					'next',
					{ type: 'headerFooterNavigate', target: 'next' },
					{ inline: true },
				),
				tool('Link to Previous', 'link', { type: 'headerFooterLink' }, { inline: true }),
			),
		),
		group(
			'Options',
			stack(
				tool(
					'Different first page',
					'firstPage',
					{ type: 'page', key: 'titlePage', value: 'toggle' },
					{ inline: true },
				),
				tool(
					'Different odd and even pages',
					'oddEven',
					{ type: 'evenOddHeaders' },
					{ inline: true },
				),
			),
		),
		group(
			'Close',
			tool('Close Header and Footer', 'close', { type: 'headerFooterClose' }, { large: true }),
		),
	);
	const position = group(
		'Position',
		stack(
			spinner(
				'Header from Top (inches)',
				(inches) => ({ type: 'headerFooterPosition', kind: 'header', inches }),
				{ min: 0, max: 22, step: 0.001 },
			),
			spinner(
				'Footer from Bottom (inches)',
				(inches) => ({ type: 'headerFooterPosition', kind: 'footer', inches }),
				{ min: 0, max: 22, step: 0.001 },
			),
		),
	);
	panel.insertBefore(position, panel.lastElementChild);
}
export function syncHeaderFooterRibbon(
	toolbar: HTMLElement,
	context: HeaderFooterContext | undefined,
	model: DocumentModel,
	editable: boolean,
): void {
	const tab = toolbar.querySelector<HTMLButtonElement>('#dve-tab-header-footer');
	if (tab) {
		const entering = tab.hidden && !!context;
		tab.hidden = !context;
		if (entering) tab.click();
		else if (!context && tab.getAttribute('aria-selected') === 'true')
			toolbar.querySelector<HTMLButtonElement>('#dve-tab-home')?.click();
	}
	for (const button of toolbar.querySelectorAll<HTMLButtonElement>('button[data-action]')) {
		const action = JSON.parse(button.dataset.action!);
		if (action.type === 'headerFooterLink') {
			button.disabled = !context || context.index === 0 || !editable;
			button.setAttribute(
				'aria-pressed',
				String(
					!!context &&
						context.index > 0 &&
						!model.sections?.[context.index]?.[context.kind]?.[context.slot],
				),
			);
		} else if (action.type === 'headerFooterClose') button.disabled = !context;
		else if (action.type === 'headerFooterNavigate')
			button.disabled =
				!context ||
				(action.target === 'previous' && context.index === 0) ||
				(action.target === 'next' && context.index >= (model.sections?.length ?? 1) - 1);
	}
	const section = context ? model.sections?.[context.index] : undefined;
	for (const [label, value] of [
		['Header from Top (inches)', section?.headerDistanceTwips ?? 720],
		['Footer from Bottom (inches)', section?.footerDistanceTwips ?? 720],
	] as const) {
		const input = toolbar.querySelector<HTMLInputElement>(`[data-localearialabel="${label}"]`);
		if (!input) continue;
		input.disabled = !context || !editable;
		const focused =
			input.getRootNode() instanceof ShadowRoot
				? (input.getRootNode() as ShadowRoot).activeElement
				: document.activeElement;
		if (focused !== input) input.value = String(+(value / 1440).toFixed(3));
	}
}
