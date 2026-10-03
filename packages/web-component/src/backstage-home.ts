import type { PageContext } from './backstage-pages';

/** A local-file start page, following the Office viewers' File navigation. */
export function renderHome({ handlers, t, content }: PageContext): void {
	const title = document.createElement('h2');
	title.textContent = t('Home');
	const name = document.createElement('h3');
	name.className = 'dve-backstage-title';
	name.textContent = handlers.summary().fileName;
	const actions = document.createElement('div');
	actions.className = 'dve-backstage-home-actions';
	for (const [label, command] of [
		['Blank document', 'new'],
		['Open', 'open'],
		['Save a copy as DOCX', 'export'],
	] as const) {
		const button = document.createElement('button');
		button.type = 'button';
		button.className = 'dve-backstage-primary';
		button.textContent = t(label);
		button.addEventListener('click', () => {
			handlers.close();
			handlers.fileCommand(command);
		});
		actions.append(button);
	}
	const note = document.createElement('p');
	note.className = 'dve-backstage-muted';
	note.textContent = t('Files are processed entirely in the browser.');
	content.replaceChildren(title, name, actions, note);
}
