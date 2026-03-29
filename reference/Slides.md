# Sequence of slides in a presentation

Supports indexed access (1-based),
[`length()`](https://rdrr.io/r/base/length.html), and iteration via
`to_list()`.

## Super classes

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\>
[`rpptx::ParentedElementProxy`](https://markheckmann.github.io/rpptx/reference/ParentedElementProxy.md)
-\> `Slides`

## Methods

### Public methods

- [`Slides$new()`](#method-Slides-new)

- [`Slides$get()`](#method-Slides-get)

- [`Slides$add_slide()`](#method-Slides-add_slide)

- [`Slides$get_by_id()`](#method-Slides-get_by_id)

- [`Slides$index()`](#method-Slides-index)

- [`Slides$to_list()`](#method-Slides-to_list)

- [`Slides$duplicate_slide()`](#method-Slides-duplicate_slide)

- [`Slides$delete()`](#method-Slides-delete)

- [`Slides$move()`](#method-Slides-move)

- [`Slides$clone()`](#method-Slides-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    Slides$new(sldIdLst, prs)

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    Slides$get(idx)

------------------------------------------------------------------------

### Method `add_slide()`

#### Usage

    Slides$add_slide(slide_layout)

------------------------------------------------------------------------

### Method `get_by_id()`

#### Usage

    Slides$get_by_id(slide_id, default = NULL)

------------------------------------------------------------------------

### Method `index()`

#### Usage

    Slides$index(slide)

------------------------------------------------------------------------

### Method `to_list()`

#### Usage

    Slides$to_list()

------------------------------------------------------------------------

### Method `duplicate_slide()`

#### Usage

    Slides$duplicate_slide(slide)

------------------------------------------------------------------------

### Method `delete()`

#### Usage

    Slides$delete(slide)

------------------------------------------------------------------------

### Method `move()`

#### Usage

    Slides$move(slide, idx)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    Slides$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
