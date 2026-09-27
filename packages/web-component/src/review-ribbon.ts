import type { ReviewDisplayMode } from './review-display';
import type { RibbonAction } from './ribbon';

function emit(control: HTMLElement, detail: RibbonAction): void {
	control.dispatchEvent(
		new CustomEvent('ribbon-action', { bubbles: true, composed: true, detail }),
	);
}
function reviewButton(
	label: string,
	text: string,
	action: RibbonAction,
	className = '',
): HTMLButtonElement {
	const el = document.createElement('button');
	el.type = 'button';
	el.textContent = text;
	el.setAttribute('aria-label', label);
	if (className) el.className = className;
	el.addEventListener('mousedown', (event) => event.preventDefault());
	el.addEventListener('click', () => emit(el, action));
	return el;
}
function group(label: string, ...children: HTMLElement[]): HTMLElement {
	const el = document.createElement('div');
	el.className = 'ribbon-group';
	el.dataset.label = label;
	el.dataset.caption = label;
	el.setAttribute('role', 'group');
	el.setAttribute('aria-label', `${label} controls`);
	el.append(...children);
	return el;
}

/** Review tab's Track Changes / navigate-and-resolve / comments controls (see AGENTS.md hub-file rule). */
export function createReviewControls(): HTMLElement[] {
	const displayModeSelect = document.createElement('select');
	displayModeSelect.setAttribute('aria-label', 'Display for review');
	for (const [value, text] of [
		['all', 'All markup'],
		['simple', 'Simple markup'],
		['final', 'No markup'],
		['original', 'Original'],
	] as const) {
		const option = document.createElement('option');
		option.value = value;
		option.textContent = text;
		displayModeSelect.append(option);
	}
	displayModeSelect.value = 'all';
	displayModeSelect.addEventListener('change', () =>
		emit(displayModeSelect, {
			type: 'reviewDisplay',
			value: displayModeSelect.value as ReviewDisplayMode,
		}),
	);

	const tracking = group(
		'Tracking',
		reviewButton(
			'Track changes',
			'Track changes',
			{ type: 'review', key: 'trackChanges' },
			'tool-track',
		),
		displayModeSelect,
	);
	const changes = group(
		'Changes',
		reviewButton('Previous change', '↑', { type: 'review', key: 'previous' }),
		reviewButton('Next change', '↓', { type: 'review', key: 'next' }),
		reviewButton('Accept', 'Accept', { type: 'review', key: 'acceptOne' }),
		reviewButton('Reject', 'Reject', { type: 'review', key: 'rejectOne' }),
		reviewButton('Accept all', 'Accept all', { type: 'review', key: 'acceptAll' }),
		reviewButton('Reject all', 'Reject all', { type: 'review', key: 'rejectAll' }),
	);
	const comments = group(
		'Comments',
		reviewButton('Add comment', 'Add comment', { type: 'comments', key: 'add' }),
		reviewButton('Comments', 'Comments', { type: 'comments', key: 'toggle' }),
	);
	return [tracking, changes, comments];
}
