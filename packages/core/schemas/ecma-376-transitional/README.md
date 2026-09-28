# ECMA-376 Transitional XML schemas

The WordprocessingML schema (`wml.xsd`) and the schemas it imports, from **ECMA-376 5th edition
(December 2016), Part 4: Transitional Migration Features**,
`OfficeOpenXML-XMLSchema-Transitional.zip`, downloaded from
https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip.
`xml.xsd` is the W3C schema for the `xml:` namespace (https://www.w3.org/2001/xml.xsd).

They are used only by tests (`src/test-support/schema-validation.ts`), which validate the parts
the writer produces against the standard. The files are unmodified; the test helper points the
schemas' `xml.xsd` imports at the local copy.

## Licence and provenance

- **ECMA-376 schemas** (every `*.xsd` except `xml.xsd`): copyright Ecma International. They are
  distributed by Ecma under the copyright notice that accompanies ECMA-376, which permits copying
  and redistributing the specification and its schema files, including for use in implementing the
  standard, provided the Ecma copyright notice is kept and the files are not misrepresented.
  The files are vendored here unmodified, solely as test fixtures for schema validation; they are
  not shipped in the published packages. The authoritative terms are those in the original
  download above; consult them before redistributing these files for any other purpose.
- **`xml.xsd`**: from the W3C (https://www.w3.org/2001/xml.xsd), distributed under the W3C
  Software and Document Notice and License (https://www.w3.org/copyright/software-license/).
  Copyright W3C (MIT, ERCIM, Keio, Beihang).

No file was edited after download. If a file must change, record it here and update its checksum.

## SHA-256 checksums

Verify with `sha256sum -c` from this directory against the list below.

```
41b93bd8857cc68b1e43be2806a872d736a9bdd6566900062d8fdb57d7bbb354  dml-chart.xsd
3fd0586f2637b98bb9886f0e0b67d89e1cc987c2d158cc7deb5f5b9890ced412  dml-chartDrawing.xsd
29b254ee0d10414a8504b5a08149c7baec35a60d5ff607d6b3f492aa36815f40  dml-diagram.xsd
5cb76dabd8b97d1e9308a1700b90c20139be4d50792d21a7f09789f5cccd6026  dml-lockedCanvas.xsd
5375417f0f5394b8dd1a7035b9679151f19a6b65df309dec10cfb4a420cb00e9  dml-main.xsd
5d389d42befbebd91945d620242347caecd3367f9a3a7cf8d97949507ae1f53c  dml-picture.xsd
bdad416b096b61d37b71603b2c949484f9070c830bdaeba93bf35e15c8900614  dml-wordprocessingDrawing.xsd
8df6a1d927bbefb091a5233bd4e872f7992f62b4065fe1f129caf4fa5da68ad7  shared-commonSimpleTypes.xsd
0d103b99a4a8652f8871552a69d42d2a3760ac6a5e3ef02d979c4273257ff6a4  shared-customXmlSchemaProperties.xsd
1e1e49c96aa5faab9d72301151044c09556182adc95c00506605ff62b24e4cb8  shared-math.xsd
12264f3c03d738311cd9237d212f1c07479e70f0cbe1ae725d29b36539aef637  shared-relationshipReference.xsd
f5ee623b08b6a66935e5aced2f5d8ad0fc71bf9e8e833cd490150c0fa94b8763  vml-main.xsd
585bedc1313b40888dcc544cb74cd939a105ee674f3b1d3aa1cc6d34f70ff155  vml-officeDrawing.xsd
133c9f64a5c5d573b78d0a474122b22506d8eadb5e063f67cdbbb8fa2f161d0e  vml-presentationDrawing.xsd
6bdeb169c3717eb01108853bd9fc5a3750fb1fa5b82abbdd854d49855a40f519  vml-spreadsheetDrawing.xsd
475dcae1e7d1ea46232db6f8481040c15e53a52a3c256831d3df204212b0e831  vml-wordprocessingDrawing.xsd
4112072b459b7a4b9fe31562223945124102185d1ad113178c4a642d11e97153  wml.xsd
61960fb3131e38022caad5360e2f33a3382578ab3c80cd58bd74320ede61b20c  xml.xsd
```
