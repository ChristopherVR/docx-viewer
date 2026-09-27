<script lang="ts">
  import { mountEditor, type EditorBinding } from './index';
  import type { DocumentModel } from '@christophervr/docx-core';
  interface Props { documentModel?: DocumentModel; readOnly?: boolean; locale?: string; ondocumentchange?: (model: DocumentModel) => void; ondocumenterror?: (error: Error) => void; }
  let { documentModel, readOnly = false, locale = 'en', ondocumentchange, ondocumenterror }: Props = $props();
  let binding: EditorBinding | undefined;
  function attach(host: HTMLElement, options: Props) {
    const normalized = (value: Props) => ({ documentModel: value.documentModel, readOnly: value.readOnly, locale: value.locale,
      onDocumentChange: value.ondocumentchange, onDocumentError: value.ondocumenterror });
    binding = mountEditor(host, normalized(options));
    return { update(next: Props) { binding?.update(normalized(next)); }, destroy() { binding?.destroy(); binding = undefined; } };
  }
  export async function load(input: Uint8Array | ArrayBuffer) { if (!binding) throw new Error('Editor is not mounted'); await binding.load(input); }
  export async function save() { if (!binding) throw new Error('Editor is not mounted'); return binding.save(); }
</script>
<div use:attach={{ documentModel, readOnly, locale, ondocumentchange, ondocumenterror }}></div>
