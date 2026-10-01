# Angular

::: warning Not published to npm yet
The Word packages are unpublished, so there is nothing to `npm install` today. Build from the repository (`bun install`, `bun run demo`). The imports below show the intended API of the `@christophervr/docx-viewer` package, with the matching framework as an application-level peer.
:::

Import the standalone component and bind its document input and change output.

```ts
import { WordEditorComponent } from '@christophervr/docx-viewer/angular';

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
