import {
	Component,
	ElementRef,
	EventEmitter,
	Input,
	Output,
	inject,
	type AfterViewInit,
	type OnChanges,
	type OnDestroy,
} from '@angular/core';
import type { DocumentModel } from '@christophervr/docx-core';
import type {
	EditorThemeMode,
	PageChangeDetail,
	RibbonActionId,
} from '@christophervr/docx-web-component';
import { eventOptions, mountEditor, pickEditorProps, type EditorBinding } from './index';
@Component({ selector: 'word-editor', standalone: true, template: '' })
export class WordEditorComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() documentModel?: DocumentModel;
	@Input() readOnly = false;
	@Input() locale = 'en';
	@Input() theme: EditorThemeMode = 'auto';
	@Input() showThumbnails = false;
	@Input() showToolbar = true;
	@Input() hiddenActions: readonly RibbonActionId[] = [];
	@Output() documentChange = new EventEmitter<DocumentModel>();
	@Output() documentError = new EventEmitter<Error>();
	@Output() pageChange = new EventEmitter<PageChangeDetail>();
	@Output() dirtyChange = new EventEmitter<boolean>();
	private host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding?: EditorBinding;
	private options() {
		return {
			...pickEditorProps(this),
			...eventOptions({
				'document-change': (model) => this.documentChange.emit(model),
				'document-error': (error) => this.documentError.emit(error),
				'page-change': (detail) => this.pageChange.emit(detail),
				'dirty-change': (dirty) => this.dirtyChange.emit(dirty),
			}),
		};
	}
	ngAfterViewInit() {
		this.binding = mountEditor(this.host.nativeElement, this.options());
	}
	ngOnChanges() {
		this.binding?.update(this.options());
	}
	ngOnDestroy() {
		this.binding?.destroy();
	}
	get element() {
		return this.binding?.element;
	}
	async load(input: Uint8Array | ArrayBuffer) {
		if (!this.binding) throw new Error('Editor is not mounted');
		await this.binding.load(input);
	}
	async save() {
		if (!this.binding) throw new Error('Editor is not mounted');
		return this.binding.save();
	}
	async download(fileName?: string) {
		if (!this.binding) throw new Error('Editor is not mounted');
		await this.binding.download(fileName);
	}
	markClean() {
		this.binding?.markClean();
	}
	get dirty() {
		return this.binding?.dirty ?? false;
	}
}
