import type { EditorView } from 'prosemirror-view';

/** Common symbols for Insert > Symbol, in Word's "recently used" spirit; language-neutral glyphs. */
export const SYMBOLS = [
	'©',
	'®',
	'™',
	'§',
	'¶',
	'†',
	'‡',
	'•',
	'…',
	'–',
	'—',
	'°',
	'±',
	'×',
	'÷',
	'≠',
	'≤',
	'≥',
	'∞',
	'€',
	'£',
	'¥',
	'←',
	'→',
	'↑',
	'↓',
	'✓',
] as const;

export type DateTimeFormat = 'long' | 'short' | 'time' | 'datetime';

/** The current moment formatted for the display locale, as Insert > Date & Time offers. */
export function formatDateTime(format: DateTimeFormat, locale: string, now = new Date()): string {
	const options: Record<DateTimeFormat, Intl.DateTimeFormatOptions> = {
		long: { dateStyle: 'long' },
		short: { dateStyle: 'short' },
		time: { timeStyle: 'short' },
		datetime: { dateStyle: 'medium', timeStyle: 'short' },
	};
	try {
		return new Intl.DateTimeFormat(locale, options[format]).format(now);
	} catch {
		return new Intl.DateTimeFormat('en', options[format]).format(now);
	}
}

/** Inserts plain text at the selection (replacing it), with the caret's stored formatting. */
export function insertPlainText(view: EditorView, text: string): boolean {
	if (!view.editable || !text) return false;
	view.dispatch(view.state.tr.insertText(text).scrollIntoView());
	return true;
}
