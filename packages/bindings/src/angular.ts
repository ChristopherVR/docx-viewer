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
import { mountEditor, type EditorBinding } from './index';
@Component({ selector: 'word-editor', standalone: true, template: '' })
export class WordEditorComponent implements AfterViewInit, OnChanges, OnDestroy {
	@Input() documentModel?: DocumentModel;
	@Input() readOnly = false;
	@Output() documentChange = new EventEmitter<DocumentModel>();
	@Output() documentError = new EventEmitter<Error>();
	private host = inject<ElementRef<HTMLElement>>(ElementRef);
	private binding?: EditorBinding;
	private options() {
		return {
			documentModel: this.documentModel,
			readOnly: this.readOnly,
			onDocumentChange: (model: DocumentModel) => this.documentChange.emit(model),
			onDocumentError: (error: Error) => this.documentError.emit(error),
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
}
