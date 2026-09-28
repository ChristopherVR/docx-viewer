<script lang="ts">
  import { eventOptions, mountEditor, pickEditorProps, type EditorBinding, type EditorEventHandlers, type EditorProps } from './index';
  type Props = EditorProps & {
    ondocumentchange?: EditorEventHandlers['document-change'];
    ondocumenterror?: EditorEventHandlers['document-error'];
    onpagechange?: EditorEventHandlers['page-change'];
    ondirtychange?: EditorEventHandlers['dirty-change'];
  };
  let { documentModel, readOnly = false, locale = 'en', theme = 'auto', showThumbnails = false, showToolbar = true, hiddenActions = [], ondocumentchange, ondocumenterror, onpagechange, ondirtychange }: Props = $props();
  let binding: EditorBinding | undefined;
  function attach(host: HTMLElement, options: Props) {
    const normalized = (value: Props) => ({ ...pickEditorProps(value),
      ...eventOptions({ 'document-change': value.ondocumentchange, 'document-error': value.ondocumenterror, 'page-change': value.onpagechange, 'dirty-change': value.ondirtychange }) });
    binding = mountEditor(host, normalized(options));
    return { update(next: Props) { binding?.update(normalized(next)); }, destroy() { binding?.destroy(); binding = undefined; } };
  }
  export async function load(input: Uint8Array | ArrayBuffer) { if (!binding) throw new Error('Editor is not mounted'); await binding.load(input); }
  export async function save() { if (!binding) throw new Error('Editor is not mounted'); return binding.save(); }
  export async function download(fileName?: string) { if (!binding) throw new Error('Editor is not mounted'); await binding.download(fileName); }
  export function markClean() { binding?.markClean(); }
  export function isDirty() { return binding?.dirty ?? false; }
</script>
<div use:attach={{ documentModel, readOnly, locale, theme, showThumbnails, showToolbar, hiddenActions, ondocumentchange, ondocumenterror, onpagechange, ondirtychange }}></div>
