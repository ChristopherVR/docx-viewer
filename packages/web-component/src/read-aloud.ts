import type { EditorView } from 'prosemirror-view';

/** Whether the browser can speak text; Read Aloud is unavailable (not faked) without it. */
export function readAloudAvailable(): boolean {
	return typeof speechSynthesis !== 'undefined' && typeof SpeechSynthesisUtterance !== 'undefined';
}

/** What Word's Read Aloud reads: the selection, or everything from the caret to the end. */
export function textToRead(view: EditorView): string {
	const { doc, selection } = view.state;
	const from = selection.from;
	const to = selection.empty ? doc.content.size : selection.to;
	return doc.textBetween(from, to, '\n', ' ').trim();
}

let speaking = false;

export function isReadingAloud(): boolean {
	return speaking;
}

/**
 * Starts reading `text` with the browser's voice, or stops when it is already reading. Returns
 * whether speech is now playing. `onChange` fires when it ends by itself.
 */
export function toggleReadAloud(text: string, onChange: (speaking: boolean) => void): boolean {
	if (!readAloudAvailable()) return false;
	if (speaking) {
		speechSynthesis.cancel();
		speaking = false;
		onChange(false);
		return false;
	}
	if (!text) return false;
	const utterance = new SpeechSynthesisUtterance(text);
	const finish = () => {
		if (!speaking) return;
		speaking = false;
		onChange(false);
	};
	utterance.addEventListener('end', finish);
	utterance.addEventListener('error', finish);
	speaking = true;
	speechSynthesis.speak(utterance);
	onChange(true);
	return true;
}
