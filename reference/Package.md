# PowerPoint package

Top-level container for a .pptx file. Extends OpcPackage with
presentation-specific behavior.

## Super class

[`rpptx::OpcPackage`](https://markheckmann.github.io/rpptx/reference/OpcPackage.md)
-\> `Package`

## Methods

### Public methods

- [`Package$next_image_partname()`](#method-Package-next_image_partname)

- [`Package$next_media_partname()`](#method-Package-next_media_partname)

- [`Package$get_or_add_media_part()`](#method-Package-get_or_add_media_part)

- [`Package$get_or_add_image_part()`](#method-Package-get_or_add_image_part)

- [`Package$iter_parts()`](#method-Package-iter_parts)

- [`Package$clone()`](#method-Package-clone)

Inherited methods

- [`rpptx::OpcPackage$drop_rel()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-drop_rel)
- [`rpptx::OpcPackage$initialize()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-initialize)
- [`rpptx::OpcPackage$iter_rels()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-iter_rels)
- [`rpptx::OpcPackage$next_partname()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-next_partname)
- [`rpptx::OpcPackage$part_related_by()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-part_related_by)
- [`rpptx::OpcPackage$relate_to()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-relate_to)
- [`rpptx::OpcPackage$related_part()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-related_part)
- [`rpptx::OpcPackage$save()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-save)
- [`rpptx::OpcPackage$target_ref()`](https://markheckmann.github.io/rpptx/reference/OpcPackage.html#method-target_ref)

------------------------------------------------------------------------

### Method `next_image_partname()`

#### Usage

    Package$next_image_partname(ext)

------------------------------------------------------------------------

### Method `next_media_partname()`

#### Usage

    Package$next_media_partname(ext)

------------------------------------------------------------------------

### Method `get_or_add_media_part()`

#### Usage

    Package$get_or_add_media_part(video)

------------------------------------------------------------------------

### Method `get_or_add_image_part()`

#### Usage

    Package$get_or_add_image_part(image_file)

------------------------------------------------------------------------

### Method `iter_parts()`

#### Usage

    Package$iter_parts()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Package$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
