import {
	mountEditor,
	type EditorOptions,
	type EditorHandle,
} from '../../packages/bindings/src/index';
import type { DocxEditorElement } from '@christophervr/docx-web-component';
/** Demo-only harness: actual framework mounts exercise each public adapter. */
export async function mountFramework(
	host: HTMLElement,
	options: EditorOptions,
	frameworkOverride?: string,
): Promise<EditorHandle> {
	const framework =
		frameworkOverride ||
		new URLSearchParams(location.search).get('framework') ||
		import.meta.env.VITE_DEMO_FRAMEWORK ||
		'vanilla';
	if (framework === 'react') {
		const [{ createRoot }, { createElement }, { WordEditor }] = await Promise.all([
			import('react-dom/client'),
			import('react'),
			import('../../packages/bindings/src/react'),
		]);
		createRoot(host).render(createElement(WordEditor, options));
	} else if (framework === 'solid') {
		const [{ render }, { createComponent }, { WordEditor }] = await Promise.all([
			import('solid-js/web'),
			import('solid-js'),
			import('../../packages/bindings/src/solid'),
		]);
		render(() => createComponent(WordEditor, options), host);
	} else if (framework === 'vue') {
		const [{ createApp, h }, { WordEditor }] = await Promise.all([
			import('vue'),
			import('../../packages/bindings/src/vue'),
		]);
		createApp({
			render: () =>
				h(WordEditor, {
					documentModel: options.documentModel,
					readOnly: options.readOnly,
					locale: options.locale,
					'onDocument-change': options.onDocumentChange,
					'onDocument-error': options.onDocumentError,
				}),
		}).mount(host);
	} else if (framework === 'svelte') {
		const [{ mount }, { default: WordEditor }] = await Promise.all([
			import('svelte'),
			import('../../packages/bindings/src/WordEditor.svelte'),
		]);
		mount(WordEditor, {
			target: host,
			props: {
				documentModel: options.documentModel,
				readOnly: options.readOnly,
				locale: options.locale,
				ondocumentchange: options.onDocumentChange,
				ondocumenterror: options.onDocumentError,
			},
		});
	} else if (framework === 'angular') {
		await import('@angular/compiler');
		const [
			{ createApplication },
			{ provideZonelessChangeDetection, createComponent },
			{ WordEditorComponent },
		] = await Promise.all([
			import('@angular/platform-browser'),
			import('@angular/core'),
			import('../../packages/bindings/src/angular'),
		]);
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const component = createComponent(WordEditorComponent, {
			environmentInjector: app.injector,
			hostElement: host,
		});
		component.setInput('documentModel', options.documentModel);
		component.setInput('readOnly', options.readOnly ?? false);
		component.setInput('locale', options.locale ?? 'en');
		component.instance.documentChange.subscribe(options.onDocumentChange);
		component.instance.documentError.subscribe(options.onDocumentError);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
	} else {
		return mountEditor(host, options);
	}
	const element = await new Promise<DocxEditorElement>((resolve, reject) => {
		const existing = host.querySelector<DocxEditorElement>('docx-editor');
		if (existing) {
			resolve(existing);
			return;
		}
		const observer = new MutationObserver(() => {
			const found = host.querySelector<DocxEditorElement>('docx-editor');
			if (found) {
				clearTimeout(timeout);
				observer.disconnect();
				resolve(found);
			}
		});
		const timeout = setTimeout(() => {
			observer.disconnect();
			reject(new Error(`${framework} editor failed to mount`));
		}, 10000);
		observer.observe(host, { childList: true, subtree: true });
	});
	return {
		element,
		load: (input) => element.load(input),
		save: () => element.save(),
		download: (fileName) => element.download(fileName),
		markClean: () => element.markClean(),
		get dirty() {
			return element.dirty;
		},
	};
}
