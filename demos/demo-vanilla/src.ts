import { initTheme } from './theme';
import './style.css';
import { createDocument } from '@christophervr/docx-core';
import type { EditorHandle } from '../../packages/bindings/src/index';
import { mountFramework } from './framework';
import { createSampleDocument } from './sample-document';

const get = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
initTheme();

const framework = new URLSearchParams(location.search).get('framework') || 'vanilla';
get('build-stamp').textContent = `docx-viewer demo · ${framework}`;

let toastTimer: ReturnType<typeof setTimeout> | undefined;
function toast(message: string) {
	const element = get('toast');
	element.textContent = message;
	element.hidden = false;
	clearTimeout(toastTimer);
	toastTimer = setTimeout(() => (element.hidden = true), 6000);
}

let editor: EditorHandle | undefined;
async function showEditor(): Promise<EditorHandle> {
	get('landing').hidden = true;
	get('workspace').hidden = false;
	editor ??= await mountFramework(get('editor'), {
		documentModel: createDocument(),
		onDocumentError(error) {
			toast(error.message);
		},
	});
	return editor;
}

async function openFile(file: File) {
	try {
		const handle = await showEditor();
		await handle.load(new Uint8Array(await file.arrayBuffer()));
		handle.element.fileName = file.name;
	} catch (error) {
		get('workspace').hidden = true;
		get('landing').hidden = false;
		const message = get('landing-error');
		message.textContent = error instanceof Error ? error.message : String(error);
		message.hidden = false;
	}
}

async function openModel(model: ReturnType<typeof createDocument>, fileName: string) {
	const handle = await showEditor();
	handle.element.documentModel = model;
	handle.element.fileName = fileName;
}

const landingFile = get<HTMLInputElement>('landing-file');
landingFile.addEventListener('change', () => {
	const file = landingFile.files?.[0];
	landingFile.value = '';
	if (file) void openFile(file);
});
get('browse').addEventListener('click', (event) => {
	event.stopPropagation();
	landingFile.click();
});
get('blank').addEventListener('click', (event) => {
	event.stopPropagation();
	void openModel(createDocument(), 'Document1.docx');
});
get('sample').addEventListener('click', (event) => {
	event.stopPropagation();
	void openModel(createSampleDocument(), 'Sample document.docx');
});
const dropzone = get('dropzone');
dropzone.addEventListener('click', () => landingFile.click());
dropzone.addEventListener('keydown', (event) => {
	if (event.target === dropzone && (event.key === 'Enter' || event.key === ' ')) {
		event.preventDefault();
		landingFile.click();
	}
});
const carriesFiles = (event: DragEvent) => Boolean(event.dataTransfer?.types.includes('Files'));
for (const type of ['dragenter', 'dragover'] as const)
	document.addEventListener(type, (event) => {
		if (!carriesFiles(event)) return;
		event.preventDefault();
		dropzone.classList.add('dragging');
	});
document.addEventListener('dragleave', (event) => {
	if (!event.relatedTarget) dropzone.classList.remove('dragging');
});
document.addEventListener('drop', (event) => {
	if (!carriesFiles(event)) return;
	event.preventDefault();
	dropzone.classList.remove('dragging');
	const file = event.dataTransfer?.files[0];
	if (file) void openFile(file);
});
