# ECMA-376 Transitional XML schemas

The WordprocessingML schema (`wml.xsd`) and the schemas it imports, from **ECMA-376 5th edition
(December 2016), Part 4: Transitional Migration Features**,
`OfficeOpenXML-XMLSchema-Transitional.zip`, downloaded from
https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip.
`xml.xsd` is the W3C schema for the `xml:` namespace (https://www.w3.org/2001/xml.xsd).

They are used only by tests (`src/test-support/schema-validation.ts`), which validate the parts
the writer produces against the standard. The files are unmodified; the test helper points the
schemas' `xml.xsd` imports at the local copy.
