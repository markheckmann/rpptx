# Slide part

Slide part

Slide part

## Super classes

[`rpptx::Part`](https://markheckmann.github.io/rpptx/reference/Part.md)
-\>
[`rpptx::XmlPart`](https://markheckmann.github.io/rpptx/reference/XmlPart.md)
-\>
[`rpptx::BaseSlidePart`](https://markheckmann.github.io/rpptx/reference/BaseSlidePart.md)
-\> `SlidePart`

## Methods

### Public methods

- [`SlidePart$get_slide()`](#method-SlidePart-get_slide)

- [`SlidePart$has_notes_slide()`](#method-SlidePart-has_notes_slide)

- [`SlidePart$add_chart_part()`](#method-SlidePart-add_chart_part)

- [`SlidePart$get_or_add_image_part()`](#method-SlidePart-get_or_add_image_part)

- [`SlidePart$get_or_add_video_media_part()`](#method-SlidePart-get_or_add_video_media_part)

- [`SlidePart$clone()`](#method-SlidePart-clone)

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

### Method `get_slide()`

#### Usage

    SlidePart$get_slide()

------------------------------------------------------------------------

### Method `has_notes_slide()`

#### Usage

    SlidePart$has_notes_slide()

------------------------------------------------------------------------

### Method `add_chart_part()`

#### Usage

    SlidePart$add_chart_part(chart_type, chart_data)

------------------------------------------------------------------------

### Method `get_or_add_image_part()`

#### Usage

    SlidePart$get_or_add_image_part(image_file)

------------------------------------------------------------------------

### Method `get_or_add_video_media_part()`

#### Usage

    SlidePart$get_or_add_video_media_part(video)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    SlidePart$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
