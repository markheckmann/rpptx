# Base class for slide parts (slides, layouts, masters)

Base class for slide parts (slides, layouts, masters)

Base class for slide parts (slides, layouts, masters)

## Super classes

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\>
[`rpptx::XmlPart`](https://markheckmann.github.io/rpptx/reference/XmlPart.md)
-\> `BaseSlidePart`

## Methods

### Public methods

- [`BaseSlidePart$clone()`](#method-BaseSlidePart-clone)

Inherited methods

- [`rpptx::Part$has_rels()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-has_rels)
- [`rpptx::Part$load_rels_from_xml()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-load_rels_from_xml)
- [`rpptx::Part$part_related_by()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-part_related_by)
- [`rpptx::Part$relate_to()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-relate_to)
- [`rpptx::Part$related_part()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-related_part)
- [`rpptx::Part$rels_xml_bytes()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-rels_xml_bytes)
- [`rpptx::Part$target_ref()`](https://markheckmann.github.io/rpptx/reference/Part.html#method-target_ref)
- [`rpptx::XmlPart$drop_rel()`](https://markheckmann.github.io/rpptx/reference/XmlPart.html#method-drop_rel)
- [`rpptx::XmlPart$initialize()`](https://markheckmann.github.io/rpptx/reference/XmlPart.html#method-initialize)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    BaseSlidePart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
