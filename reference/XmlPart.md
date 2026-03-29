# Base class for parts containing an XML payload

Provides additional methods for parsing/reserializing XML and managing
relationships. Most package parts are XmlParts.

## Super class

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\> `XmlPart`

## Methods

### Public methods

- [`XmlPart$new()`](#method-XmlPart-new)

- [`XmlPart$drop_rel()`](#method-XmlPart-drop_rel)

- [`XmlPart$clone()`](#method-XmlPart-clone)

Inherited methods

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

    XmlPart$new(partname, content_type, package, element)

------------------------------------------------------------------------

### Method `drop_rel()`

#### Usage

    XmlPart$drop_rel(rId)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    XmlPart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
