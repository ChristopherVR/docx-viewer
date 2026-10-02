# smartart.docx provenance

`smartart.docx` is copied unchanged from `ooxml-core` (`src/docx/__fixtures__/smartart.docx`),
where `scripts/docx/build-smartart-fixture.mjs` builds it from
`src/pptx/__tests__/fixtures/corpus/smartart-orgchart-assistants.pptx`: the five parts of each of
its two diagrams and the theme are copied byte for byte (only the `relId` of `dsp:dataModelExt`
is rewritten to the Word relationship id). `document.xml`, the relationships and the content types
are hand-built. It was not produced by Word, so it shows that the cached drawing is read and
shown, not that Word renders it the same way. It holds one inline and one floating (`wp:anchor`)
diagram, both with a cached `dsp:drawing`.
