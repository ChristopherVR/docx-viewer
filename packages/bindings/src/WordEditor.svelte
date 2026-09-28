<script lang="ts">
  import { eventOptions, mountEditor, pickEditorProps, type EditorBinding, type EditorEventHandlers, type EditorProps } from './index';
  type Props = EditorProps & { ondocumentchange?: EditorEventHandlers['document-change']; ondocumenterror?: EditorEventHandlers['document-error'] };
  let { documentModel, readOnly = false, locale = 'en', theme = 'auto', ondocumentchange, ondocumenterror }: Props = $props();
  let binding: EditorBinding | undefined;
  function attach(host: HTMLElement, options: Props) {
    const normalized = (value: Props) => ({ ...pickEditorProps(value),
      ...eventOptions({ 'document-change': value.ondocumentchange, 'document-error': value.ondocumenterror }) });
    binding = mountEditor(host, normalized(options));
    return { update(next: Props) { binding?.update(normalized(next)); }, destroy() { binding?.destroy(); binding = undefined; } };
  }
  export async function load(input: Uint8Array | ArrayBuffer) { if (!binding) throw new Error('Editor is not mounted'); await binding.load(input); }
  export async function save() { if (!binding) throw new Error('Editor is not mounted'); return binding.save(); }
</script>
<div use:attach={{ documentModel, readOnly, locale, theme, ondocumentchange, ondocumenterror }}></div>
