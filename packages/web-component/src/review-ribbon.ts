import type { ReviewDisplayMode } from './review-display';
import { menuSelect, tool } from './ribbon-parts';

const big = { large: true } as const;
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
	const displayMode = menuSelect(
		'Display for review',
		'markup',
		[
			['all', 'All markup'],
			['simple', 'Simple markup'],
			['final', 'No markup'],
			['original', 'Original'],
		],
		(value) => ({ type: 'reviewDisplay', value: value as ReviewDisplayMode }),
	);
	displayMode.querySelector('select')!.value = 'all';

	const tracking = group(
		'Tracking',
		tool('Track changes', 'track', { type: 'review', key: 'trackChanges' }, big),
		displayMode,
	);
	const changes = group(
		'Changes',
		tool('Accept', 'accept', { type: 'review', key: 'acceptOne' }, big),
		tool('Reject', 'reject', { type: 'review', key: 'rejectOne' }, big),
		tool('Previous change', 'previous', { type: 'review', key: 'previous' }, big),
		tool('Next change', 'next', { type: 'review', key: 'next' }, big),
		tool('Accept all', 'acceptAll', { type: 'review', key: 'acceptAll' }, big),
		tool('Reject all', 'rejectAll', { type: 'review', key: 'rejectAll' }, big),
	);
	const comments = group(
		'Comments',
		tool('Add comment', 'comment', { type: 'comments', key: 'add' }, big),
		tool('Delete comment', 'deleteComment', { type: 'comments', key: 'delete' }, big),
		tool('Previous comment', 'previous', { type: 'comments', key: 'previous' }, big),
		tool('Next comment', 'next', { type: 'comments', key: 'next' }, big),
		tool('Comments', 'comments', { type: 'comments', key: 'toggle' }, big),
	);
	return [tracking, changes, comments];
}
