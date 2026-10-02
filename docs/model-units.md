# Measurement units in the document model

The `DocumentModel` stores measurements in the units WordprocessingML uses on disk. Since the
unit brands landed, those units are part of the TypeScript types, so passing a font size where a
length is expected, or an unrounded pixel value where whole twips are required, is a compile
error instead of a corrupted file.

Brands are compile-time only. A `Twips` is a plain JavaScript number at runtime, and stays
assignable to `number`, so arithmetic, comparison and JSON serialization are unchanged.

## Brands

All brands are exported from `docx-core`.

| Type           | Unit                         | Schema type             | Constructor (strict) | Rounding constructor   |
| -------------- | ---------------------------- | ----------------------- | -------------------- | ---------------------- |
| `Twips`        | 1/20 point, non-negative     | `ST_TwipsMeasure`       | `twips(n)`           | `roundTwips(n)`        |
| `SignedTwips`  | 1/20 point, may be negative  | `ST_SignedTwipsMeasure` | `signedTwips(n)`     | `roundSignedTwips(n)`  |
| `HalfPoints`   | 1/2 point (`w:sz` font size) | `ST_HpsMeasure`         | `halfPoints(n)`      | `roundHalfPoints(n)`   |
| `EighthPoints` | 1/8 point (border width)     | `ST_EighthPointMeasure` | `eighthPoints(n)`    | `roundEighthPoints(n)` |
| `Emu`          | 1/914400 inch (DrawingML)    | `ST_Coordinate`         | `emu(n)`             | `roundEmu(n)`          |

`Twips` is a subtype of `SignedTwips`: an unsigned measure can be assigned to a signed field, never
the other way round.

**Strict constructors** throw `RangeError` unless the value is already a whole, safe integer
(and, for `twips`, non-negative). Parsers and tests use them.

**Rounding constructors** are for computed values (ruler drags, zoom math, pixel conversions).
They round half away from zero and throw `RangeError` only for `NaN` or infinite input.
`roundTwips` and `roundEighthPoints` clamp negative results to zero.

**Conversions from real-world units** round to the nearest whole target unit: `twipsFromPoints`,
`signedTwipsFromPoints`, `twipsFromPixels` (15 twips per CSS pixel), `signedTwipsFromPixels`,
`twipsFromInches`, `halfPointsFromPoints`, `eighthPointsFromPoints`, `emuFromPixels`.

**Conversions out** return plain numbers: `twipsToPoints`, `twipsToPixels`, `twipsToInches`,
`halfPointsToPoints`, `eighthPointsToPoints`, `emuToPixels`.

There are no `as Twips` style casts outside `units.ts`; use a constructor.

## Which model fields carry which brand

Field names are unchanged.

- **Paragraph** (`Paragraph`, `ParagraphFormatting`, style definitions): `spacingBeforeTwips`,
  `spacingAfterTwips`, `firstLineTwips`, `hangingTwips` are `Twips`; `lineSpacingTwips`,
  `indentLeftTwips`, `indentRightTwips`, `indentStartTwips`, `indentEndTwips` are `SignedTwips`.
  `TabStop.posTwips` is `SignedTwips`.
- **Numbering levels and list labels**: `indentLeftTwips` is `SignedTwips`; `hangingTwips` and
  `firstLineTwips` are `Twips`.
- **Sections**: `pageWidthTwips`, `pageHeightTwips`, `marginLeftTwips`, `marginRightTwips`,
  `headerDistanceTwips`, `footerDistanceTwips`, `gutterTwips` and the column
  `widthTwips`/`spacingTwips` are `Twips`; `marginTopTwips` and `marginBottomTwips` are
  `SignedTwips` (Word allows negative top and bottom margins).
- **Tables**: `Table.grid` is `Twips[]`; `Table.widthTwips`, `TableCell.widthTwips` and
  `TableRowProperties.heightTwips` are `Twips`; `Table.indentTwips` is `SignedTwips` (Word writes
  negative table indents routinely); `TableCellMargins.top/bottom/left/right` are `SignedTwips`.
- **Borders**: `TableBorderSide.sizeEighthPoints` is `EighthPoints`.
- **Runs**: `TextRun.characterSpacingTwips` is `SignedTwips`.
- **Layout input** (`@christophervr/ooxml-core/docx/layout`): the paragraph `*Twips` fields on
  `LayoutParagraph` use the same brands as the model.

### Fields that are deliberately not branded

- `TextRun.fontSize` is in **points** and may be fractional (`10.5`). `HalfPoints` is the on-disk
  `w:sz` value; the parser converts it. Use `halfPointsFromPoints` when writing a size back.
- `TableBorderSide.spacePoints` is whole **points** (`ST_PointMeasure`).
- `widthPx`, `heightPx`, `offsetXPx`, `offsetYPx` and every layout `*Px` field are **CSS pixels at
  96 dpi**. They are derived, device-space values and are frequently fractional, so they stay plain
  numbers. Image extents are stored in pixels (derived from EMU) rather than as `Emu`.
- `DocumentModel.page` (`width`, `height`, margins) is in CSS pixels; `SectionProperties` is the
  authoritative twip source.

## Breaking change and migration

This is a type-level breaking change for code that builds or edits model objects with plain
numbers. Runtime behavior of well-formed values is identical.

```ts
// Before
paragraph.spacingAfterTwips = 240;
table.grid = [3000, 3000];
side.sizeEighthPoints = 8;

// After
import { twips, eighthPoints, roundTwips, twipsFromPixels } from 'docx-core';
paragraph.spacingAfterTwips = twips(240);
table.grid = [twips(3000), twips(3000)];
side.sizeEighthPoints = eighthPoints(8);

// Computed values: round explicitly at the boundary
paragraph.indentLeftTwips = roundSignedTwips(dragPx * 15);
section.pageWidthTwips = twipsFromPixels(page.width);
```

Reading is unaffected: branded values are numbers, so `paragraph.spacingAfterTwips ?? 0` and
`twipsToPixels(x)` work as before. When you need a plain number in a branded slot from a source you
trust to be whole (for example a value read back from your own JSON), pass it through `twips()` or
`signedTwips()` so a bad value fails loudly.

Exported signature changes:

- `parseSignedTwips` now returns `SignedTwips | undefined` (was `Twips`); `parseTwips` still returns
  `Twips | undefined`.
- `twips(n)` now also rejects negative numbers; use `signedTwips(n)` for signed values.
- `TocOptions.contentWidthTwips` is `Twips`.
- `@christophervr/ooxml-core/docx/layout`: `twipsToPx` takes `SignedTwips`; `pxToTwips` returns `SignedTwips`
  (rounded); new `NO_TWIPS`; `paragraphFloats` moved to its own module but is still exported from the
  package root.

## Parsing follows the schema signedness

Because the model now distinguishes signed and unsigned lengths, the parsers apply the schema's
type per attribute. Negative `w:spacing/@before|@after`, `w:ind/@firstLine|@hanging`, page
`w:pgMar/@left|@right|@header|@footer|@gutter`, table grid, widths and row heights are rejected
(the field is left undefined), while negative `w:ind/@left|@right|@start|@end`, `w:spacing/@line`,
`w:tab/@pos`, `w:pgMar/@top|@bottom`, `w:tblInd` and `w:spacing` (character spacing) are kept.
`validateModel` (the pre-save validator) applies the same classification and rejects non-integers.

## Editor attribute boundary

ProseMirror node attributes are untyped. The editor converts them into the model in one place
(`attr-units.ts` in `web-component`): values are rounded to whole twips, and non-numeric, non-finite
or (for unsigned fields) negative values are dropped rather than written.
