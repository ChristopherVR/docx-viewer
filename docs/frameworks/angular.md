# Angular

Import the standalone component and bind its document input and change output.

```ts
import { WordEditorComponent } from '@christophervr/docx-bindings/angular';

@Component({
	standalone: true,
	imports: [WordEditorComponent],
	template: '<word-editor [documentModel]="model" (documentChange)="model = $event" />',
})
export class EditorComponent {
	model = initialDocument;
}
```

See the [complete binding contract](/bindings) or [try the Angular demo](/demo-angular/).
