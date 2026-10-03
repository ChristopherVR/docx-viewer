/** Keep keyboard navigation within the File view and restore its opening control. */
export function backstageFocus(element: HTMLElement) {
	let opener: HTMLElement | null = null;
	const siblings = new Map<HTMLElement, boolean>();
	const active = () => (element.getRootNode() as Document | ShadowRoot).activeElement;
	element.addEventListener('keydown', (event) => {
		if (event.key !== 'Tab') return;
		const controls = [
			...element.querySelectorAll<HTMLElement>(
				'button:not(:disabled), input:not(:disabled), select:not(:disabled), [tabindex="0"]',
			),
		].filter((control) => !control.closest('[hidden]'));
		const first = controls[0];
		const last = controls.at(-1);
		if (event.shiftKey && active() === first) {
			event.preventDefault();
			last?.focus();
		} else if (!event.shiftKey && active() === last) {
			event.preventDefault();
			first?.focus();
		}
	});
	return {
		open() {
			const current = active();
			if (current instanceof HTMLElement && !element.contains(current)) opener = current;
			for (const sibling of element.parentElement?.children ?? []) {
				if (!(sibling instanceof HTMLElement) || sibling === element || siblings.has(sibling))
					continue;
				siblings.set(sibling, Boolean(sibling.inert));
				sibling.inert = true;
			}
		},
		close() {
			for (const [sibling, wasInert] of siblings) sibling.inert = wasInert;
			siblings.clear();
			if (opener?.isConnected) opener.focus();
			opener = null;
		},
	};
}
