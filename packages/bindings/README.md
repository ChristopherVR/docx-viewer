# @christophervr/docx-bindings

Framework adapters for the shared DOCX editor. Install `@christophervr/docx-viewer` for the public umbrella package, then import only the framework entry you use. Framework packages are optional peers.

The Solid entry is `@christophervr/docx-bindings/solid` (or `@christophervr/docx-viewer/solid`). It exports `WordEditor`, which accepts the shared editor options plus an optional `class` and `editorRef` callback. `editorRef` receives the editor element and its `load` and `save` methods. The component mounts and destroys the shared editor with its Solid owner lifecycle.
