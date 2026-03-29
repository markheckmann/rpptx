# Presentation Part

OPC part for the presentation XML (/ppt/presentation.xml).

## Super classes

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\>
[`rpptx::XmlPart`](https://markheckmann.github.io/rpptx/reference/XmlPart.md)
-\> `PresentationPart`

## Methods

### Public methods

- [`PresentationPart$add_slide()`](#method-PresentationPart-add_slide)

- [`PresentationPart$get_presentation()`](#method-PresentationPart-get_presentation)

- [`PresentationPart$get_slide()`](#method-PresentationPart-get_slide)

- [`PresentationPart$related_slide()`](#method-PresentationPart-related_slide)

- [`PresentationPart$related_slide_master()`](#method-PresentationPart-related_slide_master)

- [`PresentationPart$save()`](#method-PresentationPart-save)

- [`PresentationPart$slide_id()`](#method-PresentationPart-slide_id)

- [`PresentationPart$duplicate_slide()`](#method-PresentationPart-duplicate_slide)

- [`PresentationPart$rename_slide_parts()`](#method-PresentationPart-rename_slide_parts)

- [`PresentationPart$clone()`](#method-PresentationPart-clone)

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

### Method `add_slide()`

#### Usage

    PresentationPart$add_slide(slide_layout)

------------------------------------------------------------------------

### Method `get_presentation()`

#### Usage

    PresentationPart$get_presentation()

------------------------------------------------------------------------

### Method `get_slide()`

#### Usage

    PresentationPart$get_slide(slide_id)

------------------------------------------------------------------------

### Method `related_slide()`

#### Usage

    PresentationPart$related_slide(rId)

------------------------------------------------------------------------

### Method `related_slide_master()`

#### Usage

    PresentationPart$related_slide_master(rId)

------------------------------------------------------------------------

### Method [`save()`](https://rdrr.io/r/base/save.html)

#### Usage

    PresentationPart$save(path)

------------------------------------------------------------------------

### Method `slide_id()`

#### Usage

    PresentationPart$slide_id(slide_part)

------------------------------------------------------------------------

### Method `duplicate_slide()`

#### Usage

    PresentationPart$duplicate_slide(slide_part)

------------------------------------------------------------------------

### Method `rename_slide_parts()`

#### Usage

    PresentationPart$rename_slide_parts(rIds)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    PresentationPart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
