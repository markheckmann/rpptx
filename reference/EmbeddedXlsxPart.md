# Embedded Excel workbook part

Stores the raw xlsx bytes for a chart's data workbook.

## Super class

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\> `EmbeddedXlsxPart`

## Methods

### Public methods

- [`EmbeddedXlsxPart$load()`](#method-EmbeddedXlsxPart-load)

- [`EmbeddedXlsxPart$clone()`](#method-EmbeddedXlsxPart-clone)

Inherited methods

- [`rpptx::Part$drop_rel()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-drop_rel)
- [`rpptx::Part$has_rels()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-has_rels)
- [`rpptx::Part$initialize()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-initialize)
- [`rpptx::Part$load_rels_from_xml()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-load_rels_from_xml)
- [`rpptx::Part$part_related_by()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-part_related_by)
- [`rpptx::Part$relate_to()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-relate_to)
- [`rpptx::Part$related_part()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-related_part)
- [`rpptx::Part$rels_xml_bytes()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-rels_xml_bytes)
- [`rpptx::Part$target_ref()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-target_ref)

------------------------------------------------------------------------

### Method [`load()`](https://rdrr.io/r/base/load.html)

#### Usage

    EmbeddedXlsxPart$load(partname, content_type, package, blob)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    EmbeddedXlsxPart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
