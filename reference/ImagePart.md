# Image OPC part

Stores the raw bytes of an image and provides dimension scaling.

## Super class

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\> `ImagePart`

## Methods

### Public methods

- [`ImagePart$new()`](#method-ImagePart-new)

- [`ImagePart$scale()`](#method-ImagePart-scale)

- [`ImagePart$clone()`](#method-ImagePart-clone)

Inherited methods

- [`rpptx::Part$drop_rel()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-drop_rel)
- [`rpptx::Part$has_rels()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-has_rels)
- [`rpptx::Part$load_rels_from_xml()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-load_rels_from_xml)
- [`rpptx::Part$part_related_by()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-part_related_by)
- [`rpptx::Part$relate_to()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-relate_to)
- [`rpptx::Part$related_part()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-related_part)
- [`rpptx::Part$rels_xml_bytes()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-rels_xml_bytes)
- [`rpptx::Part$target_ref()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-target_ref)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    ImagePart$new(partname, content_type, package, blob, filename = NULL)

------------------------------------------------------------------------

### Method [`scale()`](https://rdrr.io/r/base/scale.html)

#### Usage

    ImagePart$scale(cx, cy)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    ImagePart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
