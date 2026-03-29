# Placeholder collection for a slide master

Like SlidePlaceholders but [`get()`](https://rdrr.io/r/base/get.html)
accepts a placeholder type string (e.g. "title", "body") rather than an
integer idx.

## Super classes

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\>
[`rpptx::ParentedElementProxy`](https://markheckmann.github.io/rpptx/reference/ParentedElementProxy.md)
-\>
[`rpptx::SlidePlaceholders`](https://markheckmann.github.io/rpptx/reference/SlidePlaceholders.md)
-\> `SlideMasterPlaceholders`

## Methods

### Public methods

- [`SlideMasterPlaceholders$get()`](#method-SlideMasterPlaceholders-get)

- [`SlideMasterPlaceholders$clone()`](#method-SlideMasterPlaceholders-clone)

Inherited methods

- [`rpptx::SlidePlaceholders$get_at()`](https://markheckmann.github.io/rpptx/reference/SlidePlaceholders.html#method-get_at)
- [`rpptx::SlidePlaceholders$initialize()`](https://markheckmann.github.io/rpptx/reference/SlidePlaceholders.html#method-initialize)
- [`rpptx::SlidePlaceholders$to_list()`](https://markheckmann.github.io/rpptx/reference/SlidePlaceholders.html#method-to_list)

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    SlideMasterPlaceholders$get(ph_type)

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    SlideMasterPlaceholders$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
