# Chart XML part

Wraps a c:chartSpace XML document. Created via
[`ChartPart_new()`](https://markheckmann.github.io/rpptx/reference/ChartPart_new.md).

## Super classes

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\>
[`rpptx::XmlPart`](https://markheckmann.github.io/rpptx/reference/XmlPart.md)
-\> `ChartPart`

## Methods

### Public methods

- [`ChartPart$load()`](#method-ChartPart-load)

- [`ChartPart$clone()`](#method-ChartPart-clone)

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

### Method [`load()`](https://rdrr.io/r/base/load.html)

#### Usage

    ChartPart$load(partname, content_type, package, blob)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    ChartPart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
