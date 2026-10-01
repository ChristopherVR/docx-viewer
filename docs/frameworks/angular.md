# Angular

::: warning Not published to npm yet
Nothing has been released yet, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below are the intended API of the self-contained `@christophervr/docx-angular-viewer` package; it needs only `@christophervr/docx-core` and @angular/core next to it.
:::

Import the standalone component and bind its document input and change output.

```ts
import { WordEditorComponent } from '@christophervr/docx-angular-viewer';

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
