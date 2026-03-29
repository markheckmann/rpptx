# Placeholder collection for a slide

Provides access to placeholder shapes on a slide. Supports indexed
access by placeholder `idx` via `[[`,
[`length()`](https://rdrr.io/r/base/length.html), and `to_list()`.

## Super classes

[`rpptx::ElementProxy`](https://markheckmann.github.io/rpptx/reference/ElementProxy.md)
-\>
[`rpptx::ParentedElementProxy`](https://markheckmann.github.io/rpptx/reference/ParentedElementProxy.md)
-\> `SlidePlaceholders`

## Methods

### Public methods

- [`SlidePlaceholders$new()`](#method-SlidePlaceholders-new)

- [`SlidePlaceholders$get()`](#method-SlidePlaceholders-get)

- [`SlidePlaceholders$get_at()`](#method-SlidePlaceholders-get_at)

- [`SlidePlaceholders$to_list()`](#method-SlidePlaceholders-to_list)

- [`SlidePlaceholders$clone()`](#method-SlidePlaceholders-clone)

------------------------------------------------------------------------

### Method `new()`

#### Usage

    SlidePlaceholders$new(spTree, parent)

------------------------------------------------------------------------

### Method [`get()`](https://rdrr.io/r/base/get.html)

#### Usage

    SlidePlaceholders$get(idx)

------------------------------------------------------------------------

### Method `get_at()`

#### Usage

    SlidePlaceholders$get_at(n)

------------------------------------------------------------------------

### Method `to_list()`

#### Usage

    SlidePlaceholders$to_list()

------------------------------------------------------------------------

### Method `clone()`

The objects of this class are cloneable with this method.

#### Usage

    SlidePlaceholders$clone(deep = FALSE)

#### Arguments

- `deep`:

  Whether to make a deep clone.
